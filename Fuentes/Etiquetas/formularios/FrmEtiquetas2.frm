VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmEtiquetas2 
   Caption         =   "Mantenimiento de Etiquetas"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   10845
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
            Picture         =   "FrmEtiquetas2.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas2.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas2.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas2.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas2.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas2.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas2.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas2.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas2.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas2.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas2.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEtiquetas2.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Desactivar Usuario"
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
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6510
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   10845
      _cx             =   19129
      _cy             =   11483
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
      Appearance      =   2
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "   &Consulta   |    &Detalle    "
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
      Begin SizerOneLibCtl.TabOne TabOne2 
         Height          =   6135
         Left            =   11490
         TabIndex        =   5
         Top             =   330
         Width           =   10755
         _cx             =   18971
         _cy             =   10821
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
         BackTabColor    =   -2147483638
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "&Diseño|&Configuración"
         Align           =   0
         CurrTab         =   1
         FirstTab        =   0
         Style           =   0
         Position        =   1
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
         Separators      =   -1  'True
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   5760
            Left            =   45
            TabIndex        =   7
            Top             =   45
            Width           =   10665
            Begin VB.CommandButton cmd 
               Caption         =   "&Eliminar"
               Enabled         =   0   'False
               Height          =   330
               Index           =   6
               Left            =   9270
               TabIndex        =   23
               Top             =   1140
               Width           =   1275
            End
            Begin VB.CommandButton cmd 
               Caption         =   "&Agregar"
               Enabled         =   0   'False
               Height          =   330
               Index           =   5
               Left            =   9270
               TabIndex        =   22
               Top             =   810
               Width           =   1275
            End
            Begin VB.CommandButton cmd 
               Caption         =   "&Copiar"
               Enabled         =   0   'False
               Height          =   330
               Index           =   4
               Left            =   9270
               TabIndex        =   21
               Top             =   1480
               Width           =   1275
            End
            Begin VB.CommandButton cmd 
               Height          =   240
               Index           =   0
               Left            =   1530
               Picture         =   "FrmEtiquetas2.frx":277E
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   480
               Width           =   240
            End
            Begin VB.Frame Frame3 
               Caption         =   "[ Contenido ]"
               Height          =   3855
               Left            =   30
               TabIndex        =   8
               Top             =   1860
               Width           =   10605
               Begin VB.Frame Frame5 
                  Height          =   3570
                  Left            =   9150
                  TabIndex        =   17
                  Top             =   210
                  Width           =   1410
                  Begin VB.CommandButton cmd 
                     Caption         =   "&Eliminar"
                     Enabled         =   0   'False
                     Height          =   330
                     Index           =   2
                     Left            =   60
                     TabIndex        =   19
                     Top             =   690
                     Width           =   1275
                  End
                  Begin VB.CommandButton cmd 
                     Caption         =   "Eliminar &Todos"
                     Enabled         =   0   'False
                     Height          =   330
                     Index           =   3
                     Left            =   60
                     TabIndex        =   18
                     Top             =   1080
                     Width           =   1275
                  End
                  Begin VB.CommandButton cmd 
                     Caption         =   "&Agregar"
                     Enabled         =   0   'False
                     Height          =   330
                     Index           =   1
                     Left            =   60
                     TabIndex        =   20
                     Top             =   150
                     Width           =   1275
                  End
               End
               Begin VSFlex7Ctl.VSFlexGrid fg 
                  Height          =   3465
                  Index           =   1
                  Left            =   90
                  TabIndex        =   9
                  Top             =   300
                  Width           =   8985
                  _cx             =   15849
                  _cy             =   6112
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
                  Rows            =   2
                  Cols            =   8
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmEtiquetas2.frx":28B0
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
            Begin VSFlex7Ctl.VSFlexGrid fg 
               Height          =   1005
               Index           =   0
               Left            =   30
               TabIndex        =   12
               Top             =   810
               Width           =   9075
               _cx             =   16007
               _cy             =   1773
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
               Rows            =   2
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmEtiquetas2.frx":299C
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
            Begin VB.TextBox TxtIdItem 
               Height          =   300
               Left            =   885
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   11
               Text            =   "TxtIdItem"
               Top             =   450
               Width           =   915
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Ítem"
               Height          =   195
               Index           =   8
               Left            =   180
               TabIndex        =   15
               Top             =   510
               Width           =   300
            End
            Begin VB.Label lblItem 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lblItem"
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
               Left            =   1800
               TabIndex        =   14
               Top             =   450
               Width           =   8835
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               Caption         =   "Detalle de Etiquetas"
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
               Left            =   0
               TabIndex        =   13
               Top             =   0
               Width           =   10620
            End
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   5760
            Left            =   -11310
            TabIndex        =   6
            Top             =   45
            Width           =   10665
            Begin VSPrinter7LibCtl.VSPrinter VSPVis 
               Height          =   5640
               Left            =   60
               Negotiate       =   -1  'True
               TabIndex        =   16
               Top             =   60
               Width           =   10560
               _cx             =   18627
               _cy             =   9948
               Appearance      =   1
               BorderStyle     =   0
               Enabled         =   -1  'True
               MousePointer    =   0
               BackColor       =   -2147483643
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _ConvInfo       =   1
               AutoRTF         =   -1  'True
               Preview         =   -1  'True
               DefaultDevice   =   0   'False
               PhysicalPage    =   -1  'True
               AbortWindow     =   -1  'True
               AbortWindowPos  =   0
               AbortCaption    =   "Imprimiendo..."
               AbortTextButton =   "Cancelar"
               AbortTextDevice =   "on the %s on %s"
               AbortTextPage   =   "Ahora Imprimiendo la Pagina %d of"
               FileName        =   ""
               MarginLeft      =   1440
               MarginTop       =   1440
               MarginRight     =   1440
               MarginBottom    =   500
               MarginHeader    =   0
               MarginFooter    =   0
               IndentLeft      =   0
               IndentRight     =   0
               IndentFirst     =   0
               IndentTab       =   720
               SpaceBefore     =   0
               SpaceAfter      =   0
               LineSpacing     =   100
               Columns         =   1
               ColumnSpacing   =   180
               ShowGuides      =   1
               LargeChangeHorz =   300
               LargeChangeVert =   300
               SmallChangeHorz =   30
               SmallChangeVert =   30
               Track           =   0   'False
               ProportionalBars=   -1  'True
               Zoom            =   214.012738853503
               ZoomMode        =   3
               ZoomMax         =   400
               ZoomMin         =   10
               ZoomStep        =   25
               EmptyColor      =   -2147483636
               TextColor       =   0
               HdrColor        =   0
               BrushColor      =   0
               BrushStyle      =   0
               PenColor        =   0
               PenStyle        =   0
               PenWidth        =   0
               PageBorder      =   0
               Header          =   ""
               Footer          =   ""
               TableSep        =   "|;"
               TableBorder     =   0
               TablePen        =   0
               TablePenLR      =   0
               TablePenTB      =   0
               NavBar          =   3
               NavBarColor     =   -2147483633
               ExportFormat    =   0
               URL             =   ""
               Navigation      =   3
               NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
            End
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6135
         Left            =   45
         TabIndex        =   1
         Top             =   330
         Width           =   10755
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   5655
            Left            =   0
            TabIndex        =   3
            Top             =   450
            Width           =   10710
            _ExtentX        =   18891
            _ExtentY        =   9975
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Ítem"
            Columns(0).DataField=   "desitem"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nº Etiq."
            Columns(1).DataField=   "numetiq"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=11695"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=11615"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1826"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1746"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Named:id=33:Normal"
            _StyleDefs(45)  =   ":id=33,.parent=0"
            _StyleDefs(46)  =   "Named:id=34:Heading"
            _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(48)  =   ":id=34,.wraptext=-1"
            _StyleDefs(49)  =   "Named:id=35:Footing"
            _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(51)  =   "Named:id=36:Selected"
            _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(53)  =   "Named:id=37:Caption"
            _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(55)  =   "Named:id=38:HighlightRow"
            _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=39:EvenRow"
            _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(59)  =   "Named:id=40:OddRow"
            _StyleDefs(60)  =   ":id=40,.parent=33"
            _StyleDefs(61)  =   "Named:id=41:RecordSelector"
            _StyleDefs(62)  =   ":id=41,.parent=34"
            _StyleDefs(63)  =   "Named:id=42:FilterBar"
            _StyleDefs(64)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Etiquetas"
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
            Height          =   240
            Left            =   105
            TabIndex        =   2
            Top             =   60
            Width           =   10620
         End
      End
   End
End
Attribute VB_Name = "FrmEtiquetas2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMINGRESOALMACEN.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO PARA EL INGRESO DE DOCUMENTOS NO CONTABLES DE INGRESO O SALIDA,
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 17/09/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstEtiq As New ADODB.Recordset                    ' RECORDSET PRINCIPAL QUE CARGARA TODAS LAS OPERACIONES REGISTRADAS
Dim QueHace As Integer                               ' VARIABLE QUE INDICA EL ESTADO DEL FORMULARIO 1 = NUEVO, 2 = MODIFICA; 3 = SOLO LECTURA
Dim SeEjecuto As Boolean                             ' VARIABLE UTILIZADA PARA EJECUTAR UNA SOLA VEZ EL EVENTO ACTIVATE
Dim Agregando As Boolean                             ' VARIABLE QUE INFORMA A LOS CONTROLES FlexGrid QUE SE ESTA AGREGADO UNA FILA
Dim Mostrando As Boolean
Dim CaracteresNumericos As String                    ' ESPECIFICA LOS CARACTERES NUMERICOS QUE PODRA SOPORTAR LOS CONTROLES TextBox
Dim CaracteresNumericos2 As String, VSPVistr As String   ' ESPECIFICA LOS CARACTERES NUMERICOS QUE PODRA SOPORTAR LOS CONTROLES TextBox
Dim mIdRegistro&                                     ' identificador del registro
Dim fOrdenLista As Boolean                           ' especfica el orden de la lista de la consulta
Dim xHorIni As Date                                  ' ESPECIFICA LA HORA DE INICIO
Dim mMesActivo As Integer                            ' --indica el mes activo
Dim fCierrePeriodo As Boolean                        ' --indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)
Dim IdMenuActivo As Integer                          ' INDICA EL CODIGO DEL MENU ACTIVO
Dim cSQL As String

Const ESTADOPENDIENTE_ = 1
Const ESTADOPROCESADO_ = 2
Const ESTADOAPROBADO_ = 3
Const ESTADOANULADO_ = 4

Const IZQUIERDA_ = 6
Const CENTRO_ = 7
Const DERECHA_ = 8

Dim CORRELATIVO As Double

Dim RstEtiqDet As New ADODB.Recordset
'*****************************************************************************************************
'* Nombre           : fVerifSiExistDocum
'* Tipo             : FUNCION
'* Descripcion      : BUSCA UN NUMERO DE DOCUMENTO EN LA TABLA alm_ingreso, SI EL NUMERO DOCUMENTO
'*                    EXISTE DEVUELVE VERDADERO
'* Paranetros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                              |                   |
'* Devuelve         : Boolean
'*****************************************************************************************************
Private Sub modelarEtiqueta()
    Dim xFila As Integer
    Dim xAltoTexto As Integer
    Dim xColAdi As Integer
    Dim ANCHOETIQUETA_ As Double
    Dim ALTOETIQUETA_ As Double
    Dim A As Integer
    Dim B As Integer
    
    If Not verificarCampos Then Exit Sub
    
    VSPVis.Columns = NulosN(fg(0).TextMatrix(fg(0).Row, 2))
    VSPVis.MarginLeft = NulosN(fg(0).TextMatrix(fg(0).Row, 3))
    VSPVis.MarginRight = NulosN(fg(0).TextMatrix(fg(0).Row, 4))
    VSPVis.MarginTop = NulosN(fg(0).TextMatrix(fg(0).Row, 5))
    VSPVis.MarginBottom = NulosN(fg(0).TextMatrix(fg(0).Row, 6))
    
    ANCHOETIQUETA_ = (VSPVis.PaperWidth / VSPVis.Columns)
    ALTOETIQUETA_ = VSPVis.PaperHeight
    
    VSPVis.BrushColor = &H80000005
    VSPVis.StartDoc
    
    VSPVis.FontName = "Agency FB"
    ALTOETIQUETA_ = 5000
    
    For A = 1 To VSPVis.Columns
        For B = 1 To fg(1).Rows - 1
            ' sE APLICAN VALORES DE ESTADO
            VSPVis.TextAlign = NulosN(fg(1).TextMatrix(B, 4))
            VSPVis.FontBold = NulosN(fg(1).TextMatrix(B, 6))
            VSPVis.FontSize = NulosN(fg(1).TextMatrix(B, 5))
            VSPVis.CurrentX = NulosN(fg(1).TextMatrix(B, 3)) + (ANCHOETIQUETA_ * (A - 1))
            VSPVis.CurrentY = NulosN(fg(1).TextMatrix(B, 2))
            VSPVis.Paragraph = NulosC(fg(1).TextMatrix(B, 1))
        Next B
        If A < VSPVis.Columns Then VSPVis.NewColumn
    Next A

    VSPVis.EndDoc
End Sub

Private Function GENERAR_SQL_ID_RST(Rst As ADODB.Recordset, nDesc As String, _
                            nCampo As String, Optional nTipoIn As String = "IN", _
                            Optional fEsNumero As Boolean = True) As String
    Dim nSQL As String
    Dim k&
    nSQL = ""
    
    If Rst.RecordCount = 0 Then Exit Function Else Rst.MoveFirst
    While Not Rst.EOF
        If Trim(CStr(Rst("" & nDesc & ""))) <> "" Then
            If fEsNumero = True Then
                nSQL = nSQL & NulosN(Rst("" & nDesc & "")) & ","
            Else
                nSQL = nSQL & "'" & NulosC(Rst("" & nDesc & "")) & "',"
            End If
        End If
        Rst.MoveNext
    Wend
    
    If nSQL <> "" Then nSQL = " " & nCampo & " " & nTipoIn & " (" + Left(nSQL, Len(nSQL) - 1) & ") "
        
    GENERAR_SQL_ID_RST = nSQL
End Function

Private Sub procesarRegistro(IDETIQ_ As Integer, CORR_ As Integer, CAMPO_ As String, vALOR_ As String)
    Dim xRs As New ADODB.Recordset
    
    ' SE CONFIGURA EL RECORDSET
    procesarRST
    
    ' HALLAMOS EL CORRELATIVO
    If CORR_ = 0 Then
        RstEtiqDet.Filter = "idetiq=" & IDETIQ_
        If RstEtiqDet.RecordCount = 0 Then
            CORR_ = 1
        Else
            RstEtiqDet.Sort = "corr DESC"
            RstEtiqDet.MoveFirst
            CORR_ = RstEtiqDet("corr") + 1
        End If
        
        fg(1).TextMatrix(fg(1).Row, 7) = CORR_
    End If
    
    RstEtiqDet.Filter = "idetiq=" & IDETIQ_ & " AND corr=" & CORR_
    
    If RstEtiqDet.RecordCount = 0 Then
        RstEtiqDet.AddNew
        RstEtiqDet("idetiq") = IDETIQ_
        RstEtiqDet("corr") = CORR_
    End If
    RstEtiqDet(CAMPO_) = vALOR_
    RstEtiqDet.Update
End Sub

Private Sub procesarRST()
    Dim xRs As New ADODB.Recordset
    
    ' SE CONFIGURA EL RECORDSET
    If RstEtiqDet.State = 0 Then
        cSQL = "SELECT TOP 1 * " _
            + vbCr + "FROM mae_etiquetadet;"
        
        Set xRs = Nothing
        RST_Busq xRs, cSQL, xCon
        
        Set RstEtiqDet = Nothing
        DEFINIR_RST_TMP RstEtiqDet, xRs
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim xCampos() As String
    Dim nTitulo As String
    Dim xRs As New ADODB.Recordset
    Dim Rpta As Integer
    Dim IDETIQ_ As Integer
    Dim nSQLId As String
    Dim A As Integer
    
    If QueHace = 3 Then Exit Sub
    Select Case Index
        Case 0 ' ITEM
            ReDim xCampos(3, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Producto":    xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":      xCampos(1, 1) = "codpro":        xCampos(1, 2) = "1500":         xCampos(1, 3) = "C"
            xCampos(2, 0) = "Uni. Med":    xCampos(2, 1) = "abrev":         xCampos(2, 2) = "1200":         xCampos(2, 3) = "C"
                    
            nTitulo = "Buscando Ítems"
            
            RstEtiq.Filter = adFilterNone
            nSQLId = GENERAR_SQL_ID_RST(RstEtiq, "iditem", " AND alm_inventario.id", "NOT IN", True)
        
            cSQL = "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.stckact, alm_inventario.activo " _
                + vbCr + "FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed " _
                + vbCr + "WHERE (((alm_inventario.activo)=-1)) " & nSQLId _
                + vbCr + "ORDER BY alm_inventario.codpro"
            
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "descripcion", "descripcion", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            lblItem.Caption = NulosC(xRs("descripcion"))
            TxtIdItem.Text = NulosN(xRs("id"))
            fg(0).SetFocus
        
        Case 1 ' AGREGAR
            Agregando = True
            If fg(1).Rows = fg(1).FixedRows Then
                fg(1).Rows = fg(1).Rows + 1
            Else
                fg(1).AddItem "", fg(1).Row + 1
            End If
            fg(1).Col = 1
            fg(1).SetFocus
            Agregando = False
        
        Case 2 ' ELIMINAR
            Rpta = MsgBox("¿Esta seguro de eliminar este registro?", vbYesNo + vbQuestion + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                If RstEtiqDet.State = 0 Then Exit Sub
                RstEtiqDet.Filter = "idetiq=" & NulosN(fg(0).TextMatrix(fg(0).Row, 7)) & " AND corr=" & NulosN(fg(1).TextMatrix(fg(1).Row, 7))
                If RstEtiqDet.RecordCount = 0 Then Exit Sub
                
                RstEtiqDet.Delete
                fg(1).RemoveItem fg(1).Row
            End If
        
        Case 3 ' ELIMINAR TODOS
            Rpta = MsgBox("¿Está seguro de eliminar todos los registros?", vbYesNo + vbQuestion + vbDefaultButton1, xTitulo)
        
            If Rpta = vbYes Then
                If RstEtiqDet.State = 0 Then Exit Sub
                RstEtiqDet.Filter = "idetiq=" & NulosN(fg(0).TextMatrix(fg(0).Row, 7))
                If RstEtiqDet.RecordCount = 0 Then Exit Sub
                limpiarRST RstEtiqDet, False
                fg(1).Rows = fg(1).FixedRows
            End If
        
        Case 4 ' COPIAR DE UNA ETIQUETA PREVIA
            ReDim xCampos(3, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Ítem":         xCampos(0, 1) = "desitem":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Etiqueta":     xCampos(1, 1) = "desetiq":      xCampos(1, 2) = "4000":         xCampos(1, 3) = "C"
            xCampos(2, 0) = "Nº Columnas":  xCampos(2, 1) = "columnas":     xCampos(2, 2) = "900":          xCampos(2, 3) = "C"
                    
            nTitulo = "Buscando Etiquetas"
            
            cSQL = "SELECT mae_etiqueta.id, mae_etiqueta.iditem, mae_etiqueta.descripcion AS desetiq, alm_inventario.descripcion AS desitem, mae_etiqueta.columnas, mae_etiqueta.marghorizq, mae_etiqueta.marghorder, mae_etiqueta.margverarr, mae_etiqueta.margveraba " _
                + vbCr + "FROM mae_etiqueta LEFT JOIN alm_inventario ON mae_etiqueta.iditem = alm_inventario.id;"
            
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "desitem", "desitem", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            ' SE LLENA LA CABECERA
            fg(0).Rows = fg(0).Rows + 1
            fg(0).TextMatrix(fg(0).Rows - 1, 2) = NulosN(xRs("columnas"))
            fg(0).TextMatrix(fg(0).Rows - 1, 3) = NulosN(xRs("marghorizq"))
            fg(0).TextMatrix(fg(0).Rows - 1, 4) = NulosN(xRs("marghorder"))
            fg(0).TextMatrix(fg(0).Rows - 1, 5) = NulosN(xRs("margverarr"))
            fg(0).TextMatrix(fg(0).Rows - 1, 6) = NulosN(xRs("margveraba"))
            fg(0).TextMatrix(fg(0).Rows - 1, 7) = CORRELATIVO
            
            ' SE LLENA EL DETALLE
            cSQL = "SELECT mae_etiquetadet.* " _
                + vbCr + "From mae_etiquetadet " _
                + vbCr + "WHERE (((mae_etiquetadet.idetiq)=" & NulosN(xRs("id")) & "));"
            
            Set xRs = Nothing
            RST_Busq xRs, cSQL, xCon
            
            Dim xRsAux As New ADODB.Recordset
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            If xRsAux.State = 0 Then DEFINIR_RST_TMP xRsAux, xRs
            
            xRs.MoveFirst
            For A = 1 To xRs.RecordCount
                xRsAux.AddNew
                xRsAux("idetiq") = CORRELATIVO
                xRsAux("corr") = A
                xRsAux("descripcion") = xRs("descripcion")
                xRsAux("posx") = xRs("posx")
                xRsAux("posy") = xRs("posy")
                xRsAux("alineacion") = xRs("alineacion")
                xRsAux("tamanio") = xRs("tamanio")
                xRsAux("negrita") = xRs("negrita")
                xRsAux.Update
                
                xRs.MoveNext
            Next A
            
            procesarRST
            CARGAR_RST_TMP RstEtiqDet, xRsAux
            CORRELATIVO = CORRELATIVO + 1
            fg(0).Row = fg(0).Rows - 1
            fg(0).SetFocus
        
        Case 5 ' AGREGAR CABECERA
            fg(0).Rows = fg(0).Rows + 1
            
        Case 6 ' ELIMINAR CABECERA
            IDETIQ_ = NulosN(fg(0).TextMatrix(fg(0).Row, 7))
            
            If IDETIQ_ = 0 Then
                ' SE ELIMINA LA FILA DEL GRID
                fg(0).RemoveItem fg(0).Row
                Exit Sub
            ElseIf IDETIQ_ < 0 Then
                Rpta = MsgBox("¿Esta seguro de eliminar la etiqueta?", vbYesNo + vbQuestion + vbDefaultButton1, xTitulo)
                If Rpta = vbNo Then Exit Sub
            Else
                Rpta = MsgBox("¿Esta seguro de eliminar la etiqueta;" & vbCr & "No se podrá revertir la operación ya que el registro se encuentra grabado?", vbYesNo + vbQuestion + vbDefaultButton1, xTitulo)
                If Rpta = vbNo Then Exit Sub
                ' SE ELIMINA EL REGISTRO DE LA BD
                xCon.Execute "DELETE * FROM mae_etiquetas WHERE idetiq=" & IDETIQ_
            End If
            
            ' SE ELIMINA EL RECORDSET TEMPORAL
            RstEtiqDet.Filter = adFilterNone
            RstEtiqDet.Filter = "idetiq=" & IDETIQ_
            limpiarRST RstEtiqDet, False
            ' SE ELIMINA LA FILA DEL GRID
            fg(0).RemoveItem fg(0).Row
            
    End Select
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstEtiq
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA LA COLUMNA SELECCIONADA DEL CONTROL DataGrid Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstEtiq.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        If fCierrePeriodo = False Then Exit Sub
        Nuevo
    End If
    
    If KeyCode = 46 Then
        If fCierrePeriodo = False Then Exit Sub
        Eliminar
    End If
    
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then VerMovimientos1 IdMenuActivo, NulosN(RstEtiq("id")), xCon
End Sub

Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim CAMPO_ As String
    Dim vALOR_ As String
    Dim IDETIQ_ As Integer
    Dim CORR_ As Integer
    
    If Index = 0 Then Exit Sub
    If Agregando Then Exit Sub
    
    Select Case Col
        Case 1 ' DECRIPCION
            CAMPO_ = "descripcion"
            
        Case 2 ' POS. VER.
            CAMPO_ = "posx"
            
        Case 3 ' POS. HOR.
            CAMPO_ = "posy"
            
        Case 4 ' ALINEACION
            CAMPO_ = "alineacion"
            
        Case 5 ' TAMAÑO
            CAMPO_ = "tamanio"
            
        Case 6 ' NEGRITA
            CAMPO_ = "negrita"

    End Select
    
    vALOR_ = NulosC(fg(1).TextMatrix(Row, Col))
    IDETIQ_ = NulosN(fg(0).TextMatrix(fg(0).Row, 7))
    CORR_ = NulosN(fg(1).TextMatrix(fg(1).Row, 7))
    
    procesarRegistro IDETIQ_, CORR_, CAMPO_, vALOR_
End Sub

Private Sub fg_DblClick(Index As Integer)
    If Index <> 0 Then Exit Sub
    TabOne2.CurrTab = 0
End Sub

Private Sub fg_EnterCell(Index As Integer)
    If QueHace = 3 Then
        fg(Index).Editable = flexEDNone
        fg(Index).SelectionMode = flexSelectionByRow
        Exit Sub
    End If
    fg(Index).Editable = flexEDKbdMouse
    fg(Index).SelectionMode = flexSelectionFree
End Sub

Private Sub fg_RowColChange(Index As Integer)
    If Agregando Then Exit Sub
    If Index = 1 Then Exit Sub
    RstEtiqDet.Filter = adFilterNone
    pCargarDatos fg(0).Row, RstEtiqDet
End Sub

Private Sub Form_Activate()
    'Modificado 13/01/11 Johan Castro
    '           Eliminar

    If SeEjecuto = False Then
        SeEjecuto = True
        
        mMesActivo = xMes
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        Dim NomMes As String
        Dim Cerrado As Boolean
        '------------------------------------------------------------------------------------------
        ' bloqueamos los botones del toolbar
        CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
        '------------------------------------------------------------------------------------------
        pCargarGrid
    End If
End Sub

Private Sub iniciarCampos()
    Dim xRs As New ADODB.Recordset
    Dim CAMPOS As String
    
    SeEjecuto = False
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Mostrando = False
    
    fg(0).ColWidth(7) = 0
    fg(1).ColWidth(7) = 0
    
    fg(1).AllowUserResizing = flexResizeColumns
    fg(1).SelectionMode = flexSelectionFree
    fg(1).ForeColorSel = &H80000005
    fg(1).BackColorSel = &H80&
    
    CAMPOS = "#" & IZQUIERDA_ & ";" & "IZQUIERDA"
    CAMPOS = CAMPOS & "|#" & DERECHA_ & ";" & "DERECHA"
    CAMPOS = CAMPOS & "|#" & CENTRO_ & ";" & "CENTRO"
    fg(1).ColComboList(4) = CAMPOS
    CORRELATIVO = -666
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
    TabOne1.CurrTab = 1
    ActivaTool
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Agregando Movimiento"
    Bloquea
    Blanquea
    
    Agregando = True
    fg(0).Editable = flexEDKbdMouse
    fg(0).SelectionMode = flexSelectionFree
    fg(0).Rows = fg(0).FixedRows + 1
    fg(0).TextMatrix(fg(0).Rows - 1, 7) = CORRELATIVO
    CORRELATIVO = CORRELATIVO + 1
    
    fg(1).Editable = flexEDKbdMouse
    fg(1).Rows = 1
    fg(1).SelectionMode = flexSelectionFree
    Agregando = False
    
    xHorIni = Time
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A AJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    QueHace = 3
    iniciarCampos
End Sub

Private Sub Form_Resize()
    TabOne1.Width = Me.Width - 120
    TabOne1.Height = Me.Height - 765
    Label14.Width = TabOne1.Width - 135
    Label5.Width = TabOne1.Width - 135
    Dg1.Width = TabOne1.Width - 135
    Dg1.Height = TabOne1.Height - 850
    TabOne2.Width = TabOne1.Width - 90
    TabOne2.Height = TabOne1.Height - 375
    VSPVis.Width = TabOne2.Width - 195
    VSPVis.Height = TabOne2.Height - 495
    lblItem.Width = TabOne2.Width - 1920
    fg(0).Width = TabOne2.Width - 1680
    cmd(4).Left = TabOne2.Width - 1515
    cmd(5).Left = TabOne2.Width - 1515
    cmd(6).Left = TabOne2.Width - 1515
    Frame3.Width = TabOne2.Width - 150
    Frame3.Height = TabOne2.Height - 2280
    fg(1).Width = Frame3.Width - 1620
    fg(1).Height = Frame3.Height - 390
    Frame5.Left = Frame3.Width - 1455
    Frame5.Height = Frame3.Height - 285
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario mientras este agregando o editando un registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If RstEtiq.State = 0 Then Exit Sub
        If RstEtiq.RecordCount = 0 And QueHace <> 1 Then
            Cancel = 1
            Exit Sub
        End If
        If QueHace = 3 Then MuestraSegundoTab
    End If
End Sub

Private Sub TabOne2_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If NewTab = 0 Then
        modelarEtiqueta
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            RstEtiq.Requery
            Dg1.Refresh
            Cancelar
            
            If RstEtiq.RecordCount <> 0 Then
                RstEtiq.MoveFirst
                RstEtiq.Find "iditem=" & mIdRegistro
                If RstEtiq.EOF = True Then RstEtiq.MoveFirst
            End If
            
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        RstEtiq.Filter = ""
        TDB_FiltroLimpiar Dg1
    End If
    
    If Button.Index = 10 Then
        Unload Me
        Set RstEtiq = Nothing
    End If
End Sub

Sub Cancelar()
    QueHace = 3
    Bloquea
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    ActivaTool
End Sub

Private Function verificarCampos() As Boolean
    Dim ERROR_ As Boolean
    Dim FILA_ As Integer
    Dim COLUMNA_ As Integer
 
    ERROR_ = False
    ' VERFICAMOS LAS COLUMNAS DE LA CABECERA
    For COLUMNA_ = 1 To fg(0).Cols - 1
        If fg(0).TextMatrix(1, COLUMNA_) = "" Then ERROR_ = True: Exit For
    Next COLUMNA_
           
    If ERROR_ Then
        MsgBox "Complete los datos de la cabecera de la etiqueta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(0).SetFocus
        verificarCampos = False
        Exit Function
    End If
    
    If fg(1).Rows = fg(1).FixedRows Then
        MsgBox "No ha especificado correctamente el detalle para la Etiqueta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(1).Rows = fg(1).FixedRows + 1
        fg(1).SetFocus
        verificarCampos = False
        Exit Function
    End If
        
    ERROR_ = False
    For FILA_ = 1 To fg(1).Rows - 1
        For COLUMNA_ = 2 To fg(1).Cols - 1
            If COLUMNA_ = 6 Then GoTo SIGUIENTE_
            If fg(1).TextMatrix(FILA_, COLUMNA_) = "" Then ERROR_ = True: Exit For
SIGUIENTE_:
        Next COLUMNA_
    Next FILA_
           
    If ERROR_ Then
        MsgBox "Complete los datos del detalle de la etiqueta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(1).SetFocus
        verificarCampos = False
        Exit Function
    End If
    
    verificarCampos = True
End Function

Function Grabar() As Boolean
    Dim IDETIQ_ As Integer
    Dim DESCRIPCION_ As String
    Dim IDITEM_ As Integer
    Dim COLUMNAS_ As Double
    Dim MARGHORIZQ_ As Integer
    Dim MARGHORDER_ As Integer
    Dim MARGVERARR_ As Integer
    Dim MARGVERABA_ As Integer
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    Dim B As Integer
    Dim ERROR_ As Boolean
    Dim MOSTRARMENSAJE_ As Boolean
    
    ' VERIFICAMOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
    If TxtIdItem.Text = "" Then
        MsgBox "No ha especificado ningun Ítem para la etiqueta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdItem.SetFocus
        Exit Function
    End If
    
    If fg(0).Rows = fg(0).FixedRows Then
        MsgBox "No ha especificado correctamente la cabecera para la Etiqueta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(0).Rows = fg(0).FixedRows + 1
        fg(0).SetFocus
        Exit Function
    End If
        
    If Not verificarCampos Then Exit Function
    
    MOSTRARMENSAJE_ = False
    ' Se llenan los detalles
    For A = 1 To fg(0).Rows - 1
        IDITEM_ = NulosN(TxtIdItem.Text)
        If QueHace = 1 Then IDETIQ_ = 0 Else IDETIQ_ = NulosN(fg(0).TextMatrix(A, 7))
        If IDETIQ_ < 0 Then IDETIQ_ = 0
        DESCRIPCION_ = NulosC(fg(0).TextMatrix(A, 1))
        COLUMNAS_ = NulosN(fg(0).TextMatrix(A, 2))
        MARGHORIZQ_ = NulosN(fg(0).TextMatrix(A, 3))
        MARGHORDER_ = NulosN(fg(0).TextMatrix(A, 4))
        MARGVERARR_ = NulosN(fg(0).TextMatrix(A, 5))
        MARGVERABA_ = NulosN(fg(0).TextMatrix(A, 6))
        
        ' Se prepara el Recordset
        If xRs.State = 0 Then DEFINIR_RST_TMP xRs, RstEtiqDet
        limpiarRST xRs
        RstEtiqDet.Filter = adFilterNone
        RstEtiqDet.Filter = "idetiq=" & NulosN(fg(0).TextMatrix(A, 7))
        If RstEtiqDet.RecordCount = 0 Then Exit Function
        CARGAR_RST_TMP xRs, RstEtiqDet
        ' Se graba la Etiqueta
        If A = fg(0).Rows - 1 Then MOSTRARMENSAJE_ = True
        Grabar = grabarEtiqueta(DESCRIPCION_, IDITEM_, COLUMNAS_, MARGHORIZQ_, MARGHORDER_, _
                                        MARGVERARR_, MARGVERABA_, xRs, IDETIQ_, QueHace, MOSTRARMENSAJE_)
    Next A

    mIdRegistro = IDITEM_
End Function

Public Function grabarEtiqueta(DESCRIPCION_ As String, IDITEM_ As Integer, _
                                    COLUMNAS_ As Double, _
                                    MARGHORIZQ_ As Integer, MARGHORDER_ As Integer, _
                                    MARGVERARR_ As Integer, MARGVERABA_ As Integer, _
                                    RSTDET_ As ADODB.Recordset, Optional IDETIQ_ As Integer, _
                                    Optional QUEHACE_ As Integer, _
                                    Optional MOSTRARMENSAJE_ As Boolean = True) As Boolean
    Dim xId As Double
    Dim xIdDet As Double
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    
On Error GoTo LaCague
    
    xCon.BeginTrans
    
    If IDETIQ_ = 0 Then
        ' Obetenemos el Id del registro
        xId = HallaCodigoTabla("mae_etiqueta", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM mae_etiqueta", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = IDETIQ_
        RST_Busq RstCab, "SELECT * FROM mae_etiqueta WHERE id=" & xId, xCon
        xCon.Execute "DELETE * FROM mae_etiquetadet WHERE idetiq=" & xId
    End If
    
    RST_Busq RstDet, "SELECT TOP 1 * FROM mae_etiquetadet", xCon
    
    RstCab("descripcion") = DESCRIPCION_
    RstCab("iditem") = IDITEM_
    RstCab("columnas") = COLUMNAS_
    RstCab("marghorizq") = MARGHORIZQ_
    RstCab("marghorder") = MARGHORDER_
    RstCab("margverarr") = MARGVERARR_
    RstCab("margveraba") = MARGVERABA_
    RstCab.Update
    
    RSTDET_.MoveFirst
    While Not RSTDET_.EOF
        RstDet.AddNew
        RstDet("idetiq") = xId
        RstDet("corr") = NulosN(RSTDET_("corr"))
        RstDet("descripcion") = NulosC(RSTDET_("descripcion"))
        RstDet("posx") = NulosN(RSTDET_("posx"))
        RstDet("posy") = NulosN(RSTDET_("posy"))
        RstDet("alineacion") = NulosN(RSTDET_("alineacion"))
        RstDet("tamanio") = NulosN(RSTDET_("tamanio"))
        RstDet("negrita") = NulosN(RSTDET_("negrita"))
        RstDet.Update
        
        RSTDET_.MoveNext
    Wend
            
    'Grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, 84, QUEHACE_, Time, Time, Date, xCon, xId
   
    xCon.CommitTrans
    If MOSTRARMENSAJE_ Then MsgBox "La etiqueta se registró con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstCab = Nothing
    Set RstDet = Nothing
    grabarEtiqueta = True
    Exit Function
LaCague:
'Resume
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    grabarEtiqueta = False
End Function

Sub Eliminar()
    Dim Rpta As Integer
    
    TabOne1.CurrTab = 0
    If RstEtiq.State = 0 Then Exit Sub
    If RstEtiq.RecordCount = 0 Then
        MsgBox "No hay Registros de Ingreso/Salida de Almacén para eliminar", vbInformation, xTitulo
        Exit Sub
    End If
    
    Rpta = MsgBox("¿Esta seguro de eliminar todas las etiquetas de " + NulosC(RstEtiq("desitem")) + "?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        
        ' SE ELIMINA DETALLE
        xCon.Execute "DELETE mae_etiquetadet.* " _
            + vbCr + "FROM mae_etiqueta LEFT JOIN mae_etiquetadet ON mae_etiqueta.id = mae_etiquetadet.idetiq " _
            + vbCr + "WHERE (((mae_etiqueta.iditem)=" & NulosN(RstEtiq("iditem")) & "));"
        ' SE ELIMINA CABECERA
        xCon.Execute "DELETE * FROM mae_etiqueta WHERE iditem = " & NulosN(RstEtiq("iditem"))
        
        ' Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & NulosN(RstEtiq("iditem")) & " AND idform = " & IdMenuActivo

        RstEtiq.Requery
        Dg1.Refresh
        MsgBox "El conjunto de etiquetas se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    TabOne1.CurrTab = 1
    ActivaTool
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Modificando Movimiento"
    QueHace = 2
    Bloquea
    Blanquea
    
    Agregando = True
    fg(0).Editable = flexEDKbdMouse
    fg(0).Rows = 1
    fg(0).Rows = fg(1).Rows + 1
    fg(0).SelectionMode = flexSelectionFree
    
    fg(1).Editable = flexEDKbdMouse
    fg(1).Rows = 1
    fg(1).Rows = fg(1).Rows + 1
    fg(1).SelectionMode = flexSelectionFree
    Agregando = False
    
    MuestraSegundoTab
    xHorIni = Time
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA Y DESACTIVA LA BARRA DE HERRAMIENTAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : BLANQUEA LOS CONTROLES TextBox, PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    Dim A As Integer
    
    TxtIdItem.Text = ""
    lblItem.Caption = ""
    fg(0).Rows = fg(0).FixedRows
    fg(1).Rows = fg(1).FixedRows
End Sub

'*****************************************************************************************************
'* Nombre           : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA Y DESACTIVA LOS CONTROLES TEXTBOX, PREPARA PARA AGREGAR O MODIFICAR UN
'*                    REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Bloquea()
    TxtIdItem.Locked = Not TxtIdItem.Locked
    habilitar cmd, Not TxtIdItem.Locked
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL DETALLE DEL REGISTRO ACTUAL EN LA PESTAÑA DETALLE DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    '********************************************************************
    ' Modificado: 02/04/2012 - Jose Chacon - Modificar referencias a lote
    '********************************************************************
    Dim A As Integer
    Dim xRs As New ADODB.Recordset
    
    If RstEtiq.RecordCount = 0 Then Exit Sub
    If RstEtiq.BOF = True Or RstEtiq.EOF = True Then Exit Sub
    
    CORRELATIVO = -666
    TabOne2.CurrTab = 1
    TxtIdItem.Text = NulosN(RstEtiq("iditem"))
    lblItem.Caption = NulosC(RstEtiq("desitem"))
    
    cSQL = "SELECT * " _
        + vbCr + "FROM mae_etiqueta " _
        + vbCr + "WHERE ((mae_etiqueta.iditem) = " & NulosN(RstEtiq("iditem")) & ")"
    
    fg(0).Rows = fg(0).FixedRows
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    Agregando = True
    xRs.MoveFirst
    While Not xRs.EOF
        fg(0).Rows = fg(0).Rows + 1
        fg(0).TextMatrix(fg(0).Rows - 1, 1) = NulosC(xRs("descripcion"))
        fg(0).TextMatrix(fg(0).Rows - 1, 2) = NulosN(xRs("columnas"))
        fg(0).TextMatrix(fg(0).Rows - 1, 3) = NulosN(xRs("marghorizq"))
        fg(0).TextMatrix(fg(0).Rows - 1, 4) = NulosN(xRs("marghorder"))
        fg(0).TextMatrix(fg(0).Rows - 1, 5) = NulosN(xRs("margverarr"))
        fg(0).TextMatrix(fg(0).Rows - 1, 6) = NulosN(xRs("margveraba"))
        fg(0).TextMatrix(fg(0).Rows - 1, 7) = NulosN(xRs("id"))
        
        xRs.MoveNext
    Wend
    
    cSQL = "SELECT mae_etiquetadet.* " _
        + vbCr + "FROM mae_etiqueta LEFT JOIN mae_etiquetadet ON mae_etiqueta.id = mae_etiquetadet.idetiq " _
        + vbCr + "WHERE (((mae_etiqueta.iditem)=" & NulosN(RstEtiq("iditem")) & "));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    Set RstEtiqDet = Nothing
    If RstEtiqDet.State = 0 Then DEFINIR_RST_TMP RstEtiqDet, xRs
    CARGAR_RST_TMP RstEtiqDet, xRs
    
    Agregando = False
    fg(0).Row = 1
    pCargarDatos fg(0).Row, RstEtiqDet
End Sub

Sub pCargarGrid()
    TDB_FiltroLimpiar Dg1
    Set RstEtiq = Nothing

    cSQL = "SELECT mae_etiqueta.iditem, alm_inventario.descripcion AS desitem, Count('') AS numetiq " _
        + vbCr + "FROM mae_etiqueta LEFT JOIN alm_inventario ON mae_etiqueta.iditem = alm_inventario.id " _
        + vbCr + "GROUP BY mae_etiqueta.iditem, alm_inventario.descripcion " _
        + vbCr + "ORDER BY alm_inventario.descripcion;"
        
    RST_Busq RstEtiq, cSQL, xCon
    Set Dg1.DataSource = RstEtiq
End Sub


Sub pCargarDatos(FILA_ As Integer, xRs As ADODB.Recordset)
    fg(1).Rows = fg(1).FixedRows
    If xRs.State = 0 Then Exit Sub
    xRs.Filter = adFilterNone
    xRs.Filter = "idetiq = " & fg(0).TextMatrix(FILA_, 7)
    If xRs.RecordCount = 0 Then Exit Sub
    
    Agregando = True
    xRs.Sort = "corr"
    xRs.MoveFirst
    While Not xRs.EOF
        fg(1).Rows = fg(1).Rows + 1
        fg(1).TextMatrix(fg(1).Rows - 1, 1) = NulosC(xRs("descripcion"))
        fg(1).TextMatrix(fg(1).Rows - 1, 2) = NulosN(xRs("posx"))
        fg(1).TextMatrix(fg(1).Rows - 1, 3) = NulosN(xRs("posy"))
        fg(1).TextMatrix(fg(1).Rows - 1, 4) = NulosN(xRs("alineacion"))
        fg(1).TextMatrix(fg(1).Rows - 1, 5) = NulosN(xRs("tamanio"))
        fg(1).TextMatrix(fg(1).Rows - 1, 6) = NulosN(xRs("negrita"))
        fg(1).TextMatrix(fg(1).Rows - 1, 7) = NulosN(xRs("corr"))
        
        xRs.MoveNext
    Wend
    Agregando = False
End Sub

