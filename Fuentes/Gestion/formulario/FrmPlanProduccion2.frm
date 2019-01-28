VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmPlanProduccion2 
   Caption         =   "Produccion - Plan de Produccion"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7695
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo Formato"
      Height          =   375
      Left            =   8370
      TabIndex        =   31
      Top             =   270
      Width           =   1425
   End
   Begin VB.Frame FrmProgreso 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   1065
      Left            =   3135
      TabIndex        =   19
      Top             =   3000
      Visible         =   0   'False
      Width           =   5625
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   90
         TabIndex        =   20
         Top             =   645
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando Datos"
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
         Left            =   105
         TabIndex        =   22
         Top             =   75
         Width           =   1575
      End
      Begin VB.Label LblProcesa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando Datos"
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
         Left            =   105
         TabIndex        =   21
         Top             =   420
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000002&
         Height          =   300
         Left            =   30
         Top             =   30
         Width           =   5550
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   5610
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   1035
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5610
         X2              =   5610
         Y1              =   15
         Y2              =   1050
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   5610
         Y1              =   15
         Y2              =   15
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   4
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
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar plan de produccion"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar plan de produccion"
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
                  Text            =   "Eliminar plan de produccion"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Desactivar plan de produccion"
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
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Quitar Filtro"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Plan de produccion productos terminados"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Plan de produccion de produccion productois"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6810
         Top             =   60
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
               Picture         =   "FrmPlanProduccion2.frx":0000
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanProduccion2.frx":0544
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanProduccion2.frx":08D6
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanProduccion2.frx":0A5A
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanProduccion2.frx":0EAE
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanProduccion2.frx":0FC6
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanProduccion2.frx":150A
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanProduccion2.frx":1A4E
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanProduccion2.frx":1B62
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanProduccion2.frx":1C76
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanProduccion2.frx":20CA
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanProduccion2.frx":2236
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanProduccion2.frx":277E
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPlanProduccion2.frx":2A98
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7335
      Left            =   15
      TabIndex        =   5
      Top             =   360
      Width           =   11880
      _cx             =   20955
      _cy             =   12938
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
      Caption         =   "  &Consulta  |   &Detalle  "
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
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   6915
         Left            =   12525
         TabIndex        =   27
         Top             =   375
         Width           =   11790
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   6915
         Left            =   45
         TabIndex        =   9
         Top             =   375
         Width           =   11790
         Begin VB.CommandButton CmdVerEst 
            Caption         =   "&Ver Estacionalidad"
            Height          =   525
            Left            =   5760
            TabIndex        =   30
            Top             =   450
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.CommandButton CmdAddProd 
            Caption         =   "Agregar Plan de Produccion"
            Height          =   525
            Left            =   8940
            TabIndex        =   29
            Top             =   450
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.CommandButton CmdAdd 
            Caption         =   "Agregar Plan de Ventas"
            Height          =   525
            Left            =   7350
            TabIndex        =   28
            Top             =   450
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Frame Frame15 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame15"
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   7260
            TabIndex        =   24
            Top             =   6570
            Width           =   4365
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "= Item con Stock"
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
               Left            =   645
               TabIndex        =   26
               Top             =   45
               Width           =   1470
            End
            Begin VB.Shape Shape3 
               BackColor       =   &H00C00000&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00C0C0C0&
               Height          =   180
               Left            =   0
               Top             =   45
               Width           =   540
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "= Item sin Stock"
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
               Left            =   2940
               TabIndex        =   25
               Top             =   45
               Width           =   1395
            End
            Begin VB.Shape Shape4 
               BackColor       =   &H000000C0&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00C0C0C0&
               Height          =   180
               Left            =   2295
               Top             =   45
               Width           =   540
            End
         End
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   5835
            Left            =   30
            TabIndex        =   15
            Top             =   1065
            Width           =   11775
            _cx             =   20770
            _cy             =   10292
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
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FrontTabColor   =   -2147483633
            BackTabColor    =   12632256
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483641
            Caption         =   " Terminados  | Intermedios "
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
            Begin VB.Frame Frame4 
               BackColor       =   &H008080FF&
               BorderStyle     =   0  'None
               Height          =   5475
               Left            =   15
               TabIndex        =   17
               Top             =   15
               Width           =   11745
               Begin VSFlex7Ctl.VSFlexGrid Fg2 
                  Height          =   5310
                  Left            =   45
                  TabIndex        =   18
                  Top             =   75
                  Width           =   11655
                  _cx             =   20558
                  _cy             =   9366
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
                  Rows            =   1
                  Cols            =   19
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmPlanProduccion2.frx":2E2A
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
            Begin VB.Frame Frame3 
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   0  'None
               Height          =   5475
               Left            =   -12360
               TabIndex        =   16
               Top             =   15
               Width           =   11745
               Begin VSFlex7Ctl.VSFlexGrid Fg1 
                  Height          =   5310
                  Left            =   45
                  TabIndex        =   3
                  Top             =   75
                  Width           =   11655
                  _cx             =   20558
                  _cy             =   9366
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
                  ForeColorSel    =   16777215
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
                  Cols            =   19
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmPlanProduccion2.frx":306C
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
         Begin VB.TextBox TxtDesc 
            Height          =   300
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   0
            Text            =   "TxtDesc"
            Top             =   420
            Width           =   4680
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
            Height          =   300
            Left            =   960
            TabIndex        =   1
            Top             =   735
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
            Locked          =   -1  'True
            Valor           =   "06/02/2006"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
            Height          =   300
            Left            =   4350
            TabIndex        =   2
            Top             =   735
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
            Locked          =   -1  'True
            Valor           =   "06/02/2006"
         End
         Begin VB.Label LblNumReg 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "LblNumReg"
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
            Height          =   210
            Left            =   10530
            TabIndex        =   23
            Top             =   735
            Width           =   915
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Inicio"
            Height          =   195
            Left            =   60
            TabIndex        =   14
            Top             =   765
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Descripcion"
            Height          =   195
            Left            =   60
            TabIndex        =   13
            Top             =   450
            Width           =   840
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle Plan de Produccion"
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
            Left            =   105
            TabIndex        =   12
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Nº Registros : "
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   10485
            TabIndex        =   11
            Top             =   435
            Width           =   1020
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Termino"
            Height          =   195
            Left            =   3165
            TabIndex        =   10
            Top             =   765
            Width           =   930
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6915
         Left            =   -12435
         TabIndex        =   6
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6570
            Left            =   30
            TabIndex        =   7
            Top             =   345
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   11589
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
            Columns(1).Caption=   "Nº Proyecto"
            Columns(1).DataField=   "id"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descripcion"
            Columns(2).DataField=   "descripcion"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch. Ini"
            Columns(3).DataField=   "fchini"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fch. Fin"
            Columns(4).DataField=   "fchfin"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Estado"
            Columns(5).DataField=   "estado"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2381"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2302"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=8202"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=8123"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1826"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1746"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1799"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1720"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=1905"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1826"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HDBFDFD&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H400000&"
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(64)  =   ":id=34,.wraptext=-1"
            _StyleDefs(65)  =   "Named:id=35:Footing"
            _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(67)  =   "Named:id=36:Selected"
            _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=37:Caption"
            _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(71)  =   "Named:id=38:HighlightRow"
            _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=39:EvenRow"
            _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(75)  =   "Named:id=40:OddRow"
            _StyleDefs(76)  =   ":id=40,.parent=33"
            _StyleDefs(77)  =   "Named:id=41:RecordSelector"
            _StyleDefs(78)  =   ":id=41,.parent=34"
            _StyleDefs(79)  =   "Named:id=42:FilterBar"
            _StyleDefs(80)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Consulta Plan de Produccion"
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
            Left            =   105
            TabIndex        =   8
            Top             =   30
            Width           =   11595
         End
      End
   End
   Begin VB.Menu menu01 
      Caption         =   "menu01"
      Visible         =   0   'False
      Begin VB.Menu menu01_1 
         Caption         =   "Ver estacionalidad      "
      End
   End
End
Attribute VB_Name = "FrmPlanProduccion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMPLANPRODUCCION
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO PARA EL INGRESO Y EDICION DEL PLAN DE PRODUCCION
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 10/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstPlaPro As New ADODB.Recordset
Dim RstInter As New ADODB.Recordset
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim cSQL As String


'*****************************************************************************************************
'* Nombre Archivo   : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA E INGRESO DE UN NUEVO REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Nuevo()
    QueHace = 1
    xHorIni = Time
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Blanquea
    Bloquea
    Fg1.Rows = 1
    Fg2.Rows = 1
    Label1.Caption = "Agregando Plan de Produccion"
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.BackColorSel = &H80000005
    Fg1.ForeColorSel = &HFF&
    
    Fg2.SelectionMode = flexSelectionByRow
    Fg2.BackColorSel = &H80000005
    Fg2.ForeColorSel = &HFF&
    
    TxtDesc.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA LA EDICION DE UN REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Modificar()
    QueHace = 2
    xHorIni = Time
    Label1.Caption = "Modificando Plan de Produccion"
    'Blanquea
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    PreparaRST TxtFchIni.Valor, TxtFchFin.Valor
    
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.BackColorSel = &H80000005
    Fg1.ForeColorSel = &HFF&
    
    Fg2.SelectionMode = flexSelectionByRow
    Fg2.BackColorSel = &H80000005
    Fg2.ForeColorSel = &HFF&
    
    Fg1.Editable = flexEDKbdMouse
    Fg2.Editable = flexEDKbdMouse
    TxtDesc.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA Y DESACTIVA LOS BOTONES DE LA BARRA DE HERRAMIENTAS DEL FORMULARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
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
    
    Toolbar1.Buttons(12).Enabled = Not Toolbar1.Buttons(12).Enabled
    Toolbar1.Buttons(14).Enabled = Not Toolbar1.Buttons(14).Enabled
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA LOS CONTROLES TEXTBOX DEL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Blanquea()
    TxtDesc.Text = ""
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES TEXTBOX DEL FORMULARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Bloquea()
    TxtDesc.Locked = Not TxtDesc.Locked
    TxtFchIni.Locked = Not TxtFchIni.Locked
    TxtFchFin.Locked = Not TxtFchFin.Locked
End Sub

Private Sub CmdAdd_Click()
    ' EJECUTA LA BUSQUEDA DE UN PLAN DE VENTAS
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCodItem As String
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":       xCampos(1, 1) = "id":            xCampos(1, 2) = "2000":    xCampos(1, 3) = "N"

    xform.SQLCad = "SELECT ges_planventas.id, ges_planventas.descripcion From ges_planventas " _
        & "ORDER BY ges_planventas.id"
    
    xform.Titulo = "Buscando Plan de Ventas"
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        Dim xId As Integer
        xId = xRs("id")
        Set xform = Nothing
        Set xRs = Nothing
    
        MostrarDetallePlanVentas xId
        MostrarIntermediosProducto xId
        
        LblNumReg.Caption = Fg1.Rows - 1
        MostrarSaldo
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

'******************************************************************************************
'Nombre Sub o Function: MostrarSaldo
'Tipo:                  PROCEDIMIENTO
'Descripcion:           MUESTRA EL SALDO ACUTAL DE LOS ITEMS MOSTRADOS EN EL CONTROL Fg1
'Hecho por:             ENRIQUE POLLONGO SIERRA
'Modificado:            Por Jose Chacon:
'                           -Modificacion del Stock actual conforme al Plan de Produccion
'******************************************************************************************
Sub MostrarSaldo()
    Dim Rst As New ADODB.Recordset
    Dim RstAux As New ADODB.Recordset
    Dim RstPlaProd As New ADODB.Recordset
    Dim A As Integer
    Dim xTotal As Double
    Dim fechIni As String
    Dim fechFin As String
    
    Dim salIni As Double
    Dim prod As Double
    'Dim cSQL As String
    
    FrmProgreso.Left = 3135
    FrmProgreso.Top = 3045
    FrmProgreso.Visible = True
    
    'Se procesan productos terminados
    ProgressBar1.Max = Fg1.Rows - 1
    FrmProgreso.Refresh
    LblProcesa.Caption = "Procesando Saldo de Productos Finales"
    For A = 1 To Fg1.Rows - 1
        ProgressBar1.Value = A
        FrmProgreso.Refresh
        
        fechIni = RstPlaPro("fchini")
        fechFin = RstPlaPro("fchfin")
        
        cSQL = "SELECT pro_producciondet.iditem, Sum(pro_producciondet.cantidad) AS SumaDecantidad " _
                + vbCr + "From pro_producciondet " _
                + vbCr + "GROUP BY pro_producciondet.iditem " _
                + vbCr + "HAVING (((pro_producciondet.iditem)=" & Fg1.TextMatrix(A, 0) & "))"
        RST_Busq Rst, cSQL, xCon
        
        cSQL = "SELECT pro_producciondet.iditem, Sum(pro_producciondet.cantidad) AS SumaDecantidad" _
                + vbCr + " FROM pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro" _
                + vbCr + " WHERE (((pro_produccion.dia)>=CDate('" & "01/01/" & AnoTra & "') And (pro_produccion.dia)<CDate('" & fechIni & "')))" _
                + vbCr + " GROUP BY pro_producciondet.iditem" _
                + vbCr + " HAVING (((pro_producciondet.iditem)=" & Fg1.TextMatrix(A, 0) & "))"
        RST_Busq RstAux, cSQL, xCon
        
        xTotal = 0
        salIni = 0
        'Se llena el saldo Inicial
        salIni = SaldoActual(Fg1.TextMatrix(A, 0), "01/01/" & AnoTra, CDate(fechIni) - 1, xCon)
        Fg1.TextMatrix(A, Fg1.Cols - 4) = Format(salIni, FORMAT_MONTO)
        'Se llena lo producido hasta la fecha desde el inicio del plan
        If Not Rst.EOF Then
            If Not RstAux.EOF Then
                prod = NulosN(Rst("SumaDecantidad") - RstAux("SumaDecantidad"))
            Else
                prod = NulosN(Rst("SumaDecantidad"))
            End If
        Else
            prod = 0
        End If
        
        Fg1.TextMatrix(A, Fg1.Cols - 3) = Format(prod, FORMAT_MONTO)
        'Se llena el total producido
        xTotal = salIni + prod
        Fg1.TextMatrix(A, Fg1.Cols - 2) = Format(xTotal, FORMAT_MONTO)
        'Se llena la diferencia
        Fg1.TextMatrix(A, Fg1.Cols - 1) = Format(xTotal - NulosN(Fg1.TextMatrix(A, Fg1.Cols - 5)), FORMAT_MONTO)
        'Se configura el detalle del color de texto ROJO:Falta Producir/AZUL:Producido de Mas
        With Fg1
            If (xTotal - NulosN(Fg1.TextMatrix(A, Fg1.Cols - 5))) < 0 Then
                .Select A, Fg1.Cols - 1, A, Fg1.Cols - 1: .FillStyle = flexFillRepeat: .CellForeColor = &HFF&
            Else
                .Select A, Fg1.Cols - 1, A, Fg1.Cols - 1: .FillStyle = flexFillRepeat: .CellForeColor = &HFF0000
            End If
        End With
        Fg1.TextMatrix(A, Fg1.Cols - 1) = Abs(NulosN(Fg1.TextMatrix(A, Fg1.Cols - 1)))
    Next A
    
    With Fg1
        .Select 1, Fg1.Cols - 5, Fg1.Rows - 1, Fg1.Cols - 1: .FillStyle = flexFillRepeat: .CellBackColor = &HFEFBEB
        .Select 1, 1, 1, 1
    End With
    
    
    'Se procesan productos intermediso
    ProgressBar1.Max = Fg2.Rows - 1
    FrmProgreso.Refresh
    LblProcesa.Caption = "Procesando Saldo de Productos Intermedios"
    
    For A = 1 To Fg2.Rows - 1
        ProgressBar1.Value = A
        FrmProgreso.Refresh
        
        
        fechIni = RstPlaPro("fchini")
        fechFin = RstPlaPro("fchfin")
        
        cSQL = "SELECT pro_producciondet.iditem, Sum(pro_producciondet.cantidad) AS SumaDecantidad " _
                + vbCr + "From pro_producciondet " _
                + vbCr + "GROUP BY pro_producciondet.iditem " _
                + vbCr + "HAVING (((pro_producciondet.iditem)=" & Fg2.TextMatrix(A, 0) & "))"
                
        RST_Busq Rst, cSQL, xCon
        
        cSQL = "SELECT pro_producciondet.iditem, Sum(pro_producciondet.cantidad) AS SumaDecantidad" _
                + vbCr + " FROM pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro" _
                + vbCr + " WHERE (((pro_produccion.dia)>=CDate('" & "01/01/" & AnoTra & "') And (pro_produccion.dia)<CDate('" & fechIni & "')))" _
                + vbCr + " GROUP BY pro_producciondet.iditem" _
                + vbCr + " HAVING (((pro_producciondet.iditem)=" & Fg2.TextMatrix(A, 0) & "))"
                
        RST_Busq RstAux, cSQL, xCon
        
        xTotal = 0
        salIni = 0
        
        'Se llena el saldo Inicial
        salIni = SaldoActual(Fg2.TextMatrix(A, 0), "01/01/" & AnoTra, CDate(fechIni) - 1, xCon)
        Fg2.TextMatrix(A, Fg2.Cols - 4) = Format(salIni, FORMAT_MONTO)
        'Se llena lo producido hasta la fecha desde el inicio del plan
        If Not Rst.EOF Then
            If Not RstAux.EOF Then
                prod = NulosN(Rst("SumaDecantidad") - RstAux("SumaDecantidad"))
            Else
                prod = NulosN(Rst("SumaDecantidad"))
            End If
        Else
            prod = 0
        End If
        
        Fg2.TextMatrix(A, Fg2.Cols - 3) = Format(prod, FORMAT_MONTO)
        'Se llena el total producido
        xTotal = salIni + prod
        Fg2.TextMatrix(A, Fg2.Cols - 2) = Format(xTotal, FORMAT_MONTO)
        'Se llena la diferencia
        Fg2.TextMatrix(A, Fg2.Cols - 1) = Format(xTotal - NulosN(Fg2.TextMatrix(A, Fg2.Cols - 5)), FORMAT_MONTO)
        
        'Se configura el detalle del color de texto ROJO:Falta Producir/AZUL:Producido de Mas
        With Fg2
            If (xTotal - NulosN(Fg2.TextMatrix(A, Fg2.Cols - 5))) < 0 Then
                .Select A, Fg2.Cols - 1, A, Fg2.Cols - 1: .FillStyle = flexFillRepeat: .CellForeColor = &HFF&
            Else
                .Select A, Fg2.Cols - 1, A, Fg2.Cols - 1: .FillStyle = flexFillRepeat: .CellForeColor = &HFF0000
            End If
        End With
        Fg2.TextMatrix(A, Fg2.Cols - 1) = Abs(NulosN(Fg2.TextMatrix(A, Fg2.Cols - 1)))
    Next A
    
    With Fg2
        .Select 1, Fg2.Cols - 5, Fg2.Rows - 1, Fg2.Cols - 1: .FillStyle = flexFillRepeat: .CellBackColor = &HFEFBEB
        .Select 1, 1, 1, 1
    End With
    
    FrmProgreso.Visible = False
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : MostrarIntermediosProducto
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LOS PRODUCTOS INTERMEDIOS DE LOS PRODUCTOS MOSTRADOS EN EL CONTROL Fg1
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub MostrarIntermediosProducto(xIdPlanVentas As Integer)
    Dim RstPlanVtas As New ADODB.Recordset
    
    Dim cadena As String
    
    Dim idMesIni As Integer
    Dim idMesFin As Integer
    Dim idAñoIni As Integer
    Dim idAñoFin As Integer
    
    Dim xMes As Integer
    Dim xAño As Integer
    Dim bandera As Boolean
    Dim contador As Integer
    
    Dim xMesAux As String

    RST_Busq RstPlanVtas, "TRANSFORM First(ges_planventasdet.cantidad) AS PrimeroDecantidad " _
        + vbCr + "SELECT ges_planventasdet.idpv, ges_planventasdet.codpro, alm_inventario.codpro AS codigo, alm_inventario.descripcion, mae_unidades.abrev " _
        + vbCr + "FROM (ges_planventasdet LEFT JOIN alm_inventario ON ges_planventasdet.codpro = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        + vbCr + "Where (((ges_planventasdet.idpv) = " & xIdPlanVentas & ")) " _
        + vbCr + "GROUP BY ges_planventasdet.idpv, ges_planventasdet.codpro, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev " _
        + vbCr + "PIVOT ges_planventasdet.idmes;", xCon
    
    
    idMesIni = Format(CDate(TxtFchIni.Valor), "m")
    idAñoIni = Format(CDate(TxtFchIni.Valor), "yyyy")
    
    xMes = idMesIni
    xAño = idAñoIni
    
    Dim indicador As Integer
    indicador = calcularIndicador(CDate(TxtFchIni.Valor), CDate(TxtFchFin.Valor))
        
    Dim RstRec As New ADODB.Recordset
    Dim A, B, C, xCol, xFil As Integer
    
    RellenarMeses Fg2, Format(TxtFchIni.Valor, "dd/mm/yyyy"), Format(TxtFchFin.Valor, "dd/mm/yyyy")
    
On Error GoTo LaCague
    
    LblProcesa.Caption = "Procesando Productos Intermedios"
    FrmProgreso.Visible = True
    ProgressBar1.Max = Fg1.Rows - 1
    
    RstInter.ActiveConnection = Nothing
    RstPlanVtas.MoveFirst
    
    For A = 1 To RstPlanVtas.RecordCount - 1
    
        ProgressBar1.Value = A
        FrmProgreso.Refresh
        
        RST_Busq RstRec, "SELECT alm_inventario.id , alm_inventario.codpro, alm_inventario.descripcion, alm_inventario.tippro, pro_receta.iditem, mae_unidades.abrev, " _
            & " pro_recetains.canpro, pro_receta.prirec FROM ((pro_receta LEFT JOIN pro_recetains ON pro_receta.id = pro_recetains.idrec) LEFT JOIN alm_inventario " _
            & " ON pro_recetains.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id WHERE (((alm_inventario.tippro)=3) " _
            & " AND ((pro_receta.iditem)=" & RstPlanVtas("codpro") & ") AND ((pro_receta.prirec)=1))", xCon
        
        If RstRec.RecordCount <> 0 Then
        
            RstRec.MoveFirst
            For B = 1 To RstRec.RecordCount
                If RstInter.RecordCount <> 0 Then
                    RstInter.MoveFirst
                    RstInter.Find "idpro = '" & RstRec("id") & "'"
                End If
                If RstInter.EOF = False Then
                    xMesAux = xMes
                    For C = 1 To indicador
                        RstInter(CStr(xMesAux)) = RstInter(CStr(xMesAux)) + (NulosN(RstPlanVtas(CStr(xMesAux))) * RstRec("canpro"))
                        xMesAux = xMesAux + 1
                        If xMesAux > 12 Then xMesAux = 1
                    Next C
                Else
                    RstInter.AddNew
                    
                    RstInter("idpro") = RstRec("id")
                    RstInter("cod_item") = RstRec("codpro")
                    RstInter("descripcion") = RstRec("Descripcion")
                    RstInter("unimed") = RstRec("abrev")
                    
                    
                    xMesAux = xMes
                    For C = 1 To indicador
                        RstInter(CStr(xMesAux)) = (NulosN(RstPlanVtas(CStr(xMesAux))) * RstRec("canpro"))
                        xMesAux = xMesAux + 1
                        If xMesAux > 12 Then xMesAux = 1
                    Next C
                End If
                
                RstRec.MoveNext
                If RstRec.EOF = True Then
                    Exit For
                End If
            Next B
        End If
        RstPlanVtas.MoveNext
    Next A
    
    Dim xMes1, xMes2, xMes3, xMes4, xMes5, xMes6, xMes7, xMes8, xMes9, xMes10, xMes11, xMes12 As Double
    Dim xCod_Item, xDescripcion, xUniMed As String
    
    RstInter.MoveFirst
    
    LblProcesa.Caption = "Reprocesando Productos Intermedios"
    FrmProgreso.Visible = True
    ProgressBar1.Max = 1000
    
    For A = 1 To 1000
        RstInter.Filter = "ope <> 1"
        
        ProgressBar1.Value = A
        FrmProgreso.Refresh
        
        If RstInter.RecordCount <> 0 Then
           
            RST_Busq RstRec, "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.codpro, alm_inventario.descripcion, alm_inventario.tippro, pro_receta.iditem, mae_unidades.abrev, " _
                & " pro_recetains.canpro, pro_receta.prirec FROM ((pro_receta LEFT JOIN pro_recetains ON pro_receta.id = pro_recetains.idrec) LEFT JOIN alm_inventario " _
                & " ON pro_recetains.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id WHERE (((alm_inventario.tippro)=3) " _
                & " AND ((pro_receta.iditem)=" & RstInter("idpro") & ") AND ((pro_receta.prirec)=1))", xCon
            
            xCod_Item = RstInter("cod_item")
            xDescripcion = RstInter("descripcion")
            xUniMed = RstInter("unimed")
            RstInter("ope") = 1
            
            If RstRec.RecordCount <> 0 Then
                RstRec.MoveFirst
                For B = 1 To RstRec.RecordCount
                    RstInter.Filter = adFilterNone
                    RstInter.MoveFirst
                    RstInter.Find "idpro = '" & RstRec("id") & "'"
                    
                    If RstInter.EOF = False Then
                    
                        xMesAux = xMes
                        For C = 1 To indicador
                            RstInter(CStr(xMesAux)) = RstInter(CStr(xMesAux)) + (RstInter(CStr(xMesAux)) * RstRec("canpro"))
                            xMesAux = xMesAux + 1
                            If xMesAux > 12 Then xMesAux = 1
                        Next C
                    Else
                        RstInter.AddNew
                        RstInter("idpro") = RstRec("id")
                        RstInter("cod_item") = RstRec("codpro")
                        RstInter("descripcion") = RstRec("descripcion")
                        RstInter("unimed") = RstRec("abrev")
                        Dim Variable As Integer
                        xMesAux = xMes
                        For C = 1 To indicador
                            Variable = RstInter(CStr(xMesAux))
                            RstInter(CStr(xMesAux)) = RstInter(CStr(xMesAux)) * RstRec("canpro")
                            
                            xMesAux = xMesAux + 1
                            If xMesAux > 12 Then xMesAux = 1
                        Next C
                    End If
                    
                    RstRec.MoveNext
                    If RstRec.EOF = True Then
                        Exit For
                    End If
                Next B
            End If
        Else
            Exit For
        End If
        RstInter.Filter = adFilterNone
    Next A
    Fg2.Cols = Fg1.Cols
    Fg2.Rows = 1
    RstInter.Filter = adFilterNone
    RstInter.MoveFirst
    RstInter.Sort = "descripcion"
    Dim xTotal As Double
    
    For A = 1 To RstInter.RecordCount
    
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(A, 0) = RstInter("idpro")
        Fg2.TextMatrix(A, 1) = RstInter("descripcion")
        Fg2.TextMatrix(A, 2) = RstInter("cod_item")
        Fg2.TextMatrix(A, 3) = RstInter("unimed")
        
        xMesAux = xMes
        xTotal = 0
        For B = 1 To indicador
            Fg2.TextMatrix(A, B + 3) = Format(RstInter("" & xMesAux & ""), FORMAT_MONTO)
            xTotal = xTotal + NulosN(RstInter("" & xMesAux & ""))
            xMesAux = xMesAux + 1
            If xMesAux > 12 Then xMesAux = 1
        Next B
        
        Fg2.TextMatrix(A, B + 3) = Format(xTotal, FORMAT_MONTO)
        
        RstInter.MoveNext
        If RstInter.EOF = True Then
            Exit For
        End If
    Next A
    ProcesaEstacionalidad xIdPlanVentas
    Exit Sub

LaCague:
    Resume
    MsgBox Err.Description
    MsgBox cadena
    MsgBox Err.Source
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : ProcesaEstacionalidad
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : DISTRIBUYE LA NECESIDAD DE LA MATERIA PRIMA EN FUNCION A LA ESTACIONALIDAD DE LA
'*                    MATERIA PRIMA
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub ProcesaEstacionalidad(xIdPlanVentas As Integer)
    Dim A, B, NumTotMes As Integer
    Dim xIdMP As Integer
    
    Dim Rst As New ADODB.Recordset
    Dim RstPlaVent As New ADODB.Recordset
    
    Dim xTotal, TotMesEst As Double
    Dim indicador As Integer
    
    Dim idMesIni As Integer
    Dim idMesFin As Integer
    Dim idAñoIni As Integer
    Dim idAñoFin As Integer
    
    Dim xMes As Integer
    Dim xMesAux As String
    Dim xAño As Integer
    Dim bandera As Boolean
    Dim contador As Integer
    
    RST_Busq RstPlaVent, "SELECT ges_planventas.id, ges_planventas.descripcion, ges_planventas.fchini, ges_planventas.fchfin, ges_planventas.activo " _
        + vbCr + "From ges_planventas " _
        + vbCr + "WHERE (((ges_planventas.id)= " & xIdPlanVentas & "));", xCon
    
    
    idMesIni = Format(RstPlaVent("fchini"), "m")
    idAñoIni = Format(RstPlaVent("fchini"), "yyyy")
    xMes = idMesIni
    xAño = idAñoIni
    
    indicador = calcularIndicador(RstPlaVent("fchini"), RstPlaVent("fchfin"))
   
    LblProcesa.Caption = "Procesando Estacionalidad"
    ProgressBar1.Max = Fg2.Rows - 1
    
    Dim cadena As String
    Dim detMeses(1 To 12) As Integer
    
On Error GoTo LaCague
    
    For A = 1 To Fg2.Rows - 1
        ProgressBar1.Value = A
        FrmProgreso.Refresh
        
        cSQL = "SELECT alm_inventario.codpro, alm_inventario.descripcion,  alm_inventario.id, alm_inventario.idmatpri" _
            + vbCr + "From alm_inventario " _
            + vbCr + "WHERE (((alm_inventario.id)=" & NulosN(Fg2.TextMatrix(A, 0)) & "))"
        
        RST_Busq Rst, cSQL, xCon
        
        cadena = "AA-" + Rst.Source
        If Rst.RecordCount <> 0 Then
            xIdMP = NulosN(Rst("idmatpri"))
            
            cSQL = "SELECT pro_estacionalidad.* From pro_estacionalidad " _
                + vbCr + "WHERE (((pro_estacionalidad.id) = " & xIdMP & "));"
            
            RST_Busq Rst, cSQL, xCon
            
            cadena = "BB-" + Rst.Source
            If Rst.RecordCount <> 0 Then
                
                detMeses(1) = Rst("ene")
                detMeses(2) = Rst("feb")
                detMeses(3) = Rst("mar")
                detMeses(4) = Rst("abr")
                detMeses(5) = Rst("may")
                detMeses(6) = Rst("jun")
                detMeses(7) = Rst("jul")
                detMeses(8) = Rst("ago")
                detMeses(9) = Rst("set")
                detMeses(10) = Rst("oct")
                detMeses(11) = Rst("nov")
                detMeses(12) = Rst("dic")
                
                NumTotMes = 0
                TotMesEst = 0
                xTotal = NulosN(Fg2.TextMatrix(A, Fg2.Cols - 5))  'cargamos el total del producto
                
                xMesAux = xMes
                For B = 1 To indicador
                    If detMeses(xMesAux) = 2 Then NumTotMes = NumTotMes + 1
                    xMesAux = xMesAux + 1
                    If xMesAux > 12 Then xMesAux = 1
                Next B

                If NumTotMes = 0 Then NumTotMes = 1
                TotMesEst = xTotal / NumTotMes
                
                xMesAux = xMes
                For B = 1 To indicador
                    If detMeses(xMesAux) = 2 Then
                        Fg2.TextMatrix(A, B + 3) = Format(TotMesEst, FORMAT_MONTO)
                    Else
                        Fg2.TextMatrix(A, B + 3) = Format(0, FORMAT_MONTO)
                    End If
                    xMesAux = xMesAux + 1
                    If xMesAux > 12 Then xMesAux = 1
                Next B
            End If
        End If
    Next A
    Exit Sub

LaCague:
    MsgBox Err.Description
End Sub


'*****************************************************************************************************
'* Nombre           : PreparaRST
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL DETALLE DEL PEDIDO PENDIENTE
'* Creado por       : ENRIQUE POLLONGO SIERRA
'* Modificado       : 30/12/10 por JOSE CHACON MANRIQUE
'                       -Cambio total en el diseño modificandose para poder procesar
'*                       no solo fechas anuales
'*****************************************************************************************************

Sub PreparaRST(Fini As String, fFin As String)
    Dim idMesIni As Integer
    Dim idMesFin As Integer
    Dim idAñoIni As Integer
    Dim idAñoFin As Integer
    Dim xFun As New eps_librerias.FuncionesData
    Dim bandera As Boolean
    bandera = True
    Dim xMes As Integer
    Dim xAño As Integer
    Dim contador As Integer
    Dim indicador As Integer
    Dim xCampos() As String
    
    bandera = True
    
    idMesIni = Format(CDate(Fini), "m")
    idAñoIni = Format(CDate(Fini), "yyyy")
    
    xMes = idMesIni
    xAño = idAñoIni
    
    indicador = calcularIndicador(CDate(Fini), CDate(fFin))
    contador = 1
    xMes = idMesIni
    xAño = idAñoIni
    
    ReDim xCampos(5 + indicador, 3) As String
    
    xCampos(0, 0) = "cod_item":     xCampos(0, 1) = "C":      xCampos(0, 2) = "20"
    xCampos(1, 0) = "unimed":       xCampos(1, 1) = "C":      xCampos(1, 2) = "4"
    xCampos(2, 0) = "descripcion":  xCampos(2, 1) = "C":      xCampos(2, 2) = "200"
    
    While bandera
        xCampos(contador + 2, 0) = "" & xMes & ""
        xCampos(contador + 2, 1) = "N":
        xCampos(contador + 2, 2) = "2"
        xMes = xMes + 1
        If xMes > 12 Then xMes = 1: xAño = xAño + 1
        contador = contador + 1
        If contador > indicador Then bandera = False
    Wend
    xCampos(contador + 2, 0) = "ope":         xCampos(contador + 2, 1) = "N":      xCampos(contador + 2, 2) = "2"
    contador = contador + 1
    xCampos(contador + 2, 0) = "idpro":       xCampos(contador + 2, 1) = "N":      xCampos(contador + 2, 2) = "2"
    Set RstInter = xFun.CrearRstTMP(xCampos)
    RstInter.Open
End Sub

'*****************************************************************************************************
'* Nombre           : MostrarDetallePlanVentas
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL DETALLE DEL PEDIDO PENDIENTE
'* Creado por       : ENRIQUE POLLONGO SIERRA
'* Modificado       :30/12/10 por JOSE CHACON MANRIQUE
'*                      -Cambio total en el diseño modificandose las consultas para
'*                       utilizacion de tablas Cruzadas
'*                      -Cambio total en el diseño modificandose para poder procesar
'*                       no solo fechas anuales
'*****************************************************************************************************
Sub MostrarDetallePlanVentas(xIdPlanVentas As Integer)
    Dim RstPlaVent As New ADODB.Recordset
    Dim RstDeta As New ADODB.Recordset
    Dim RstProd As New ADODB.Recordset
    
    Dim cadena As String
    
    On Error GoTo LaCague
    
    Dim A, B, xCol As Integer
    Dim Total As Double
    
    Dim idMesIni As Integer
    Dim idMesFin As Integer
    Dim idAñoIni As Integer
    Dim idAñoFin As Integer
    
    Dim xMes As Integer
    Dim xAño As Integer
    Dim bandera As Boolean
    Dim contador As Integer
    
    Dim xMesAux As String
    
    'Hallamos los detalles del plan de ventas actual
    RST_Busq RstPlaVent, "SELECT ges_planventas.id, ges_planventas.descripcion, ges_planventas.fchini, ges_planventas.fchfin, ges_planventas.activo " _
        + vbCr + "From ges_planventas " _
        + vbCr + "WHERE (((ges_planventas.id)= " & xIdPlanVentas & "));", xCon
        
    'Hallamos los detalles del plan de ventas
    RST_Busq RstProd, "TRANSFORM First(ges_planventasdet.cantidad) AS PrimeroDecantidad " _
        + vbCr + "SELECT ges_planventasdet.idpv, ges_planventasdet.codpro, alm_inventario.codpro AS codigo, alm_inventario.descripcion, mae_unidades.abrev " _
        + vbCr + "FROM (ges_planventasdet LEFT JOIN alm_inventario ON ges_planventasdet.codpro = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        + vbCr + "Where (((ges_planventasdet.idpv) = " & xIdPlanVentas & ")) " _
        + vbCr + "GROUP BY ges_planventasdet.idpv, ges_planventasdet.codpro, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev " _
        + vbCr + "PIVOT ges_planventasdet.idmes;", xCon
        
    
    Fg1.Rows = 1
    Fg1.Cols = 3
    
    TxtDesc.Text = RstPlaVent("descripcion")
    TxtFchIni.Valor = Format(RstPlaVent("fchini"), "dd/mm/yyyy")
    TxtFchFin.Valor = Format(RstPlaVent("fchfin"), "dd/mm/yyyy")
    
    idMesIni = Format(CDate(RstPlaVent("fchini")), "m")
    'idMesFin = CInt(Mid(Format(RstPlaVent("fchfin"), "dd/mm/yyyy"), 4, 2))
    idAñoIni = Format(CDate(RstPlaVent("fchini")), "yyyy")
    'idAñoFin = CInt(Mid(Format(RstPlaVent("fchfin"), "dd/mm/yyyy"), 7, 4))
    
    xMes = idMesIni
    xAño = idAñoIni
    
    Dim indicador As Integer
    indicador = calcularIndicador(RstPlaVent("fchini"), RstPlaVent("fchfin"))
    
    Fg1.Cols = Fg1.Cols + indicador + 4
    
    RellenarMeses Fg1, Format(RstPlaVent("fchini"), "dd/mm/yyyy"), Format(RstPlaVent("fchfin"), "dd/mm/yyyy")
    
    PreparaRST Format(RstPlaVent("fchini"), "dd/mm/yyyy"), Format(RstPlaVent("fchfin"), "dd/mm/yyyy")
    
    TabOne2.CurrTab = 0
    
    cadena = RstDeta.Source
    
    If RstProd.RecordCount <> 0 Then
        RstProd.MoveFirst
        For A = 1 To RstProd.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Total = 0
            Fg1.TextMatrix(A, 0) = NulosC(RstProd("codpro"))
            Fg1.TextMatrix(A, 1) = NulosC(RstProd("descripcion"))
            Fg1.TextMatrix(A, 2) = NulosC(RstProd("codigo"))
            Fg1.TextMatrix(A, 3) = NulosC(RstProd("abrev"))
            
            xMesAux = xMes
            For B = 1 To indicador
                Fg1.TextMatrix(A, B + 3) = Format(NulosN(RstProd("" & xMesAux & "")), FORMAT_MONTO)
                Total = Total + NulosN(RstProd("" & xMesAux & ""))
                xMesAux = xMesAux + 1
                If xMesAux > 12 Then xMesAux = 1
            Next B
            
            Fg1.TextMatrix(A, B + 3) = Format(Total, FORMAT_MONTO)
            RstProd.MoveNext
        Next A
        Set RstProd = Nothing
    End If
    Exit Sub

LaCague:
    MsgBox Err.Description
    MsgBox cadena
End Sub

Private Function calcularIndicador(fchIni As String, fchFin As String) As Integer
    Dim indicador As Integer
    Dim idMesIni As Integer
    Dim idMesFin As Integer
    Dim idAñoIni As Integer
    Dim idAñoFin As Integer
    
    idMesIni = NulosN(Format(fchIni, "m"))
    idMesFin = NulosN(Format(fchFin, "m"))
    idAñoIni = NulosN(Format(fchIni, "yyyy"))
    idAñoFin = NulosN(Format(fchFin, "yyyy"))
    
    If idMesIni <> 0 And idAñoIni <> 0 Then
        If idAñoFin > idAñoIni Then
            indicador = (13 - idMesIni) + idMesFin
        Else
            indicador = idMesFin - idMesIni + 1
        End If
        
        If indicador > 12 Then indicador = 12
    End If
    
    calcularIndicador = indicador
End Function

Sub MostrarDetallePlanProduccion(xIdPlanProduccion As Integer, detPlan As String, fchIni As String, fchFin As String)
    Dim Rst As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    Dim Rst2Aux As New ADODB.Recordset
    Dim A, B, xCol As Integer
    Dim Total As Double
    
    Dim idMesIni As Integer
    Dim idMesFin As Integer
    Dim idAñoIni As Integer
    Dim idAñoFin As Integer
    
    Dim xMes As Integer
    Dim xAño As Integer
    Dim bandera As Boolean
    Dim contador As Integer
    
    Dim xMesAux As String
    
    Fg1.Rows = 1
    Fg1.Cols = 3
    
    Fg2.Rows = 1
    Fg2.Cols = 3
    
    TxtDesc.Text = detPlan
    TxtFchIni.Valor = Format(CDate(fchIni), "dd/mm/yyyy")
    TxtFchFin.Valor = Format(CDate(fchFin), "dd/mm/yyyy")
    
    idMesIni = Format(CDate(fchIni), "m")
    idAñoIni = Format(CDate(fchIni), "yyyy")
    
    xMes = idMesIni
    xAño = idAñoIni
    
    Dim indicador As Integer
    indicador = calcularIndicador(TxtFchIni.Valor, TxtFchFin.Valor)
'
    Fg1.Cols = Fg1.Cols + indicador + 4
    
    RellenarMeses Fg1, TxtFchIni.Valor, TxtFchFin.Valor
    TabOne2.CurrTab = 0
    
    'MOSTRAMOS LOS PRODUCTOS FINALES
    RST_Busq Rst, "TRANSFORM First(ges_plaproddet.cantidad) AS PrimeroDecantidad " _
        & "SELECT ges_plaproddet.codpro, alm_inventario.codpro AS codigo, alm_inventario.descripcion, mae_unidades.abrev, ges_plaproddet.idpv, alm_inventario.stckini " _
        & "FROM (ges_plaproddet INNER JOIN alm_inventario ON ges_plaproddet.codpro = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        & "Where (((ges_plaproddet.idpv) = " & xIdPlanProduccion & ") And ((ges_plaproddet.idmes) <> 13)) " _
        & "GROUP BY ges_plaproddet.codpro, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ges_plaproddet.idpv, alm_inventario.stckini " _
        & "PIVOT ges_plaproddet.idmes", xCon

    Fg1.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Total = 0
            Fg1.TextMatrix(A, 0) = NulosC(Rst("codpro"))
            Fg1.TextMatrix(A, 1) = NulosC(Rst("descripcion"))
            Fg1.TextMatrix(A, 2) = NulosC(Rst("codigo"))
            Fg1.TextMatrix(A, 3) = NulosC(Rst("abrev"))
            
            xMesAux = xMes
            For B = 1 To indicador
                Fg1.TextMatrix(A, B + 3) = Format(NulosN(Rst("" & xMesAux & "")), FORMAT_MONTO)
                Total = Total + NulosN(Rst("" & xMesAux & ""))
                xMesAux = xMesAux + 1
                If xMesAux > 12 Then xMesAux = 1
            Next B
            
            Fg1.TextMatrix(A, B + 3) = Format(Total, FORMAT_MONTO)
            Rst.MoveNext
        Next A
        Set Rst = Nothing
    End If
    
    'MOSTRAMOS LOS PRODUCTOS INTERMEDIOS
    
    Fg2.Cols = Fg2.Cols + indicador + 4
    RellenarMeses Fg2, TxtFchIni.Valor, TxtFchFin.Valor

    RST_Busq Rst, "TRANSFORM First(ges_plaproddet2.cantidad) AS PrimeroDecantidad " _
            & "SELECT ges_plaproddet2.codpro, alm_inventario.codpro AS codigo, alm_inventario.descripcion, mae_unidades.abrev, ges_plaproddet2.idpv, alm_inventario.stckini " _
            & "FROM (ges_plaproddet2 INNER JOIN alm_inventario ON ges_plaproddet2.codpro = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
            & "Where (((ges_plaproddet2.idpv) = " & RstPlaPro("id") & ") And ((ges_plaproddet2.idmes) <> 13)) " _
            & "GROUP BY ges_plaproddet2.codpro, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ges_plaproddet2.idpv, alm_inventario.stckini " _
            & "PIVOT ges_plaproddet2.idmes", xCon

    Fg2.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Total = 0
            Fg2.TextMatrix(A, 0) = NulosC(Rst("codpro"))
            Fg2.TextMatrix(A, 1) = NulosC(Rst("descripcion"))
            Fg2.TextMatrix(A, 2) = NulosC(Rst("codigo"))
            Fg2.TextMatrix(A, 3) = NulosC(Rst("abrev"))
            
            xMesAux = xMes
            For B = 1 To indicador
                Fg2.TextMatrix(A, B + 3) = Format(NulosN(Rst("" & xMesAux & "")), FORMAT_MONTO)
                Total = Total + NulosN(Rst("" & xMesAux & ""))
                xMesAux = xMesAux + 1
                If xMesAux > 12 Then xMesAux = 1
                
            Next B
            Fg2.TextMatrix(A, B + 3) = Format(Total, FORMAT_MONTO)
            Rst.MoveNext
        Next A
    End If
    
    LblNumReg.Caption = Fg1.Rows - 1
    MostrarSaldo
    
    Fg1.Editable = flexEDKbdMouse
End Sub

Private Sub CmdAddProd_Click()
    'Se Busca un plan de Produccion
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCodItem As String
    'Dim cSQL As String
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":       xCampos(1, 1) = "id":            xCampos(1, 2) = "2000":    xCampos(1, 3) = "N"

    cSQL = "SELECT ges_plaprod.id, ges_plaprod.descripcion, ges_plaprod.fchini, ges_plaprod.fchfin " _
        + vbCr + "FROM ges_plaprod;"

    xform.SQLCad = cSQL
    
    xform.Titulo = "Buscando Plan de Produccion"
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        Dim xId As Integer
        xId = xRs("id")
        
        MostrarDetallePlanProduccion xRs("id"), xRs("descripcion"), xRs("fchini"), xRs("fchfin")
        
        Set xform = Nothing
        Set xRs = Nothing
        
        LblNumReg.Caption = Fg1.Rows - 1
        'Modificar
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdVerEst_Click()
    ' MUESTRA LA ESTACIONALIDAD DEL ITEM CARGADO EN EL CONTROL Fg2, PARA ELLO LLAMA AL FORMULARIO FrmVistaEstacionalidad
    If Fg2.Rows = 1 Then
        MsgBox "No se ha procesado ningun plan de ventas", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    If TabOne2.CurrTab = 0 Then
        FrmVistaEstacionalidad.TxtNumGrid.Text = 1
    Else
        FrmVistaEstacionalidad.TxtNumGrid.Text = 2
    End If
    FrmVistaEstacionalidad.Show
End Sub

Private Sub Command1_Click()
    FrmPlanProduccion3.Show
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstPlaPro("id")), xCon
    End If
End Sub

Private Sub Fg2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu menu01
    End If
End Sub

Private Sub Form_Activate()
'Modificado: 08/01/11 Johan Castro
'            Agregar linea de codigo para bloquear accesos de usuarios

    ' SEGUNDO EVENTO DEL FORMULARIO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    
    If SeEjecuto = False Then
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon

        '----------------------------------------------
        
        RST_Busq RstPlaPro, "SELECT ges_plaprod.*, IIf([ges_plaprod]![activo]=0,'No Activo','Activo') AS estado " _
            & " From ges_plaprod ORDER BY ges_plaprod.id DESC", xCon
        
        Set Dg1.DataSource = RstPlaPro

    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO DEL FORMULARIO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Frame3.BackColor = &H8000000F
    Frame4.BackColor = &H8000000F
    Frame15.BackColor = &H8000000F
    
    Fg1.AllowUserResizing = flexResizeColumns
    Fg1.AutoSearch = flexSearchFromTop
    Fg1.ExplorerBar = flexExSortShowAndMove
    
    Fg2.AllowUserResizing = flexResizeColumns
    Fg2.AutoSearch = flexSearchFromTop
    Fg2.ExplorerBar = flexExSortShowAndMove
    
    Fg1.ColWidth(0) = 0
    Fg1.ColWidth(2) = 0
    
    Fg2.ColWidth(0) = 0
    Fg2.ColWidth(2) = 0
    
    QueHace = 3
    SeEjecuto = False
    TabOne1.CurrTab = 0
    TabOne2.CurrTab = 0
    Fg1.SelectionMode = flexSelectionByRow
    Fg2.SelectionMode = flexSelectionByRow
    
    Fg1.FrozenCols = 3
    Fg2.FrozenCols = 3
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Grabar
'* Tipo             : FUNCION
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA ges_plaprod, ESTA FUNCION DEVUELVE VERDADERO CUANDO
'*                    EXITO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Function Grabar() As Boolean
    If NulosC(TxtDesc.Text) = "" Then
        MsgBox "No ha especificado la descripcion del plan de produccion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtDesc.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtFchIni.Valor) = "" Then
        MsgBox "No ha especificado la fecha de inicio del plan de produccion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtFchFin.Valor) = "" Then
        MsgBox "No ha especificado la fecha final del plan de produccion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Function
    End If
    
    If Fg1.Rows = 1 Then
        MsgBox "No ha procesado ningun plan de ventas para el plan de produccion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        CmdAdd.SetFocus
        Exit Function
    End If
    
    On Error GoTo LaCague
    xCon.BeginTrans
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstDet2 As New ADODB.Recordset
    Dim RstFue As New ADODB.Recordset
    Dim xId As Double
    
    Dim idMesIni As Integer
    
    Dim xMes As Integer
    
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT * FROM ges_plaprod", xCon
        
        xId = HallaCodigoTabla("ges_plaprod", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstPlaPro("id")
        
        RST_Busq RstCab, "SELECT * FROM ges_plaprod WHERE id=" & xId & " ", xCon
        xCon.Execute "DELETE * FROM ges_plaproddet WHERE idpv = " & xId & ""
        xCon.Execute "DELETE * FROM ges_plaproddet2 WHERE idpv = " & xId & ""

    End If
    
    RST_Busq RstDet, "SELECT * FROM ges_plaproddet", xCon
    RST_Busq RstDet2, "SELECT * FROM ges_plaproddet2", xCon
    
    RstCab("descripcion") = TxtDesc.Text
    RstCab("fchini") = NulosC(TxtFchIni.Valor)
    RstCab("fchfin") = NulosC(TxtFchFin.Valor)
    RstCab.Update
    
    Dim xFila, xCol As Integer
    
    idMesIni = CInt(Mid(TxtFchIni.Valor, 4, 2))
    
    
    For xFila = 1 To Fg1.Rows - 1
        xMes = idMesIni
        For xCol = 4 To Fg1.Cols - 5
            RstDet.AddNew
            RstDet("idpv") = xId
            RstDet("codpro") = Trim(Fg1.TextMatrix(xFila, 0))
            RstDet("idmes") = xMes
            RstDet("cantidad") = NulosN(Fg1.TextMatrix(xFila, xCol))
            RstDet.Update
            xMes = xMes + 1
            If xMes > 12 Then xMes = 1
        Next xCol
    Next xFila
    
    For xFila = 1 To Fg2.Rows - 1
        xMes = idMesIni
        For xCol = 4 To Fg2.Cols - 5
            RstDet2.AddNew
            RstDet2("idpv") = xId
            RstDet2("codpro") = Trim(Fg2.TextMatrix(xFila, 0))
            RstDet2("idmes") = xMes
            RstDet2("cantidad") = NulosN(Fg2.TextMatrix(xFila, xCol))
            RstDet2.Update
            xMes = xMes + 1
            If xMes > 12 Then xMes = 1
        Next xCol
    Next xFila
       
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    xCon.CommitTrans
    MsgBox "El plan de produccion se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    Grabar = True
    Exit Function

LaCague:
    'Resume
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

'*****************************************************************************************************
'* Nombre Archivo   : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL INGRESO O EDICION DE UN REGISTRO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    Label1.Caption = "Detalle Plan de Produccion"
    Bloquea
    ActivaTool
    TabOne1.TabEnabled(0) = True

    Fg1.Editable = flexEDNone
    Fg1.BackColorSel = &H80&
    Fg1.ForeColorSel = &H80000005

    Fg2.Editable = flexEDNone
    Fg2.BackColorSel = &H80&
    Fg2.ForeColorSel = &H80000005
    TabOne1.CurrTab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    End If
End Sub


Private Sub menu01_1_Click()
    CmdVerEst_Click
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then
            MuestraSegundoTab
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA ges_plaprod
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    Rpta = MsgBox("¿Esta seguro de eliminar el plan de produccion seleccionado?", vbYesNo + vbQuestion + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        
        xCon.Execute "DELETE * FROM ges_plaproddet2 WHERE idpv = " & RstPlaPro("id") & ""
        xCon.Execute "DELETE * FROM ges_plaproddet WHERE idpv = " & RstPlaPro("id") & ""
        xCon.Execute "DELETE * FROM ges_plaprod WHERE id = " & RstPlaPro("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstPlaPro("id") & " AND idform = " & IdMenuActivo
        
        RstPlaPro.Requery
        Dg1.Refresh
        
    End If
End Sub

Private Sub TabOne2_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        LblNumReg.Caption = Fg2.Rows - 1
    Else
        LblNumReg.Caption = Fg1.Rows - 1
    End If
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : CambiarEstado
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CAMBIA EL ESTADO DE UN REGISTRO EN LA TABLA ges_plaprod
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Activado     |  Boolean   |  INDICA SI SE ACTIVA O DESACTIVA EL REGISTRO
'* DEVUELVE         :
'*****************************************************************************************************
Sub CambiarEstado(Activado As Boolean)
    Dim Rpta As Integer
    TabOne1.CurrTab = 0
    If Activado = False Then
        Rpta = MsgBox("Esta seguro de desactivar el plan de abastecimiento seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    Else
        Rpta = MsgBox("Esta seguro de activar el plan de abastecimiento seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    End If
    
    If Rpta = vbYes Then
        If Activado = False Then
            xCon.Execute "UPDATE ges_plaprod SET ges_plaprod.activo = 0 Where (((ges_plaprod.id) = " & RstPlaPro("id") & "))"
            MsgBox "El plan de produccion se desactivo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Else
            xCon.Execute "UPDATE ges_plaprod SET ges_plaprod.activo = -1 Where (((ges_plaprod.id) = " & RstPlaPro("id") & "))"
            MsgBox "El plan de produccion se activo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
    End If
    RstPlaPro.Requery
    Dg1.Refresh
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo: CmdAddProd.Visible = True: CmdAdd.Visible = True: CmdVerEst.Visible = True
    
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 8 Then
        If TabOne2.CurrTab = 0 Then
            If Fg1.Rows = 1 Then
                MsgBox "No se ha procesado el plan de produccion para los productos Terminados", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
        End If
        If TabOne2.CurrTab = 1 Then
            If Fg1.Rows = 1 Then
                MsgBox "No se ha procesado el plan de produccion para los productos Intermedios", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
        End If
        ExportarExcel
    End If
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstPlaPro.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar: CmdAddProd.Visible = False: CmdAdd.Visible = False: CmdVerEst.Visible = False
    
    If Button.Index = 15 Then
        Set RstPlaPro = Nothing
        Unload Me
    End If
End Sub

Sub RellenarMeses(fgx As VSFlexGrid, Fini As String, fFin As String)
    Dim Rst As New ADODB.Recordset
    Dim bandera As Boolean
    Dim xMes As Integer
    Dim xAño As Integer
    Dim contador As Integer
    Dim indicador As Integer
    Dim idMesIni As Integer
    Dim idMesFin As Integer
    Dim idAñoIni As Integer
    Dim idAñoFin As Integer
    
    bandera = True
    
    idMesIni = Format(CDate(Fini), "m")
    'idMesFin = CInt(Mid(fFin, 4, 2))
    idAñoIni = Format(CDate(Fini), "yyyy")
    'idAñoFin = CInt(Mid(fFin, 7, 4))
    
    xMes = idMesIni
    xAño = idAñoIni
    
    indicador = calcularIndicador(CDate(Fini), CDate(fFin))
    If indicador > 12 Then indicador = 12
    contador = 1
    fgx.TextMatrix(0, contador) = "Producto"
    fgx.TextMatrix(0, contador + 2) = "Unidad"
    While bandera
        RST_Busq Rst, "SELECT DISTINCT con_meses.id, con_meses.descripcion " _
                    & "FROM con_meses " _
                    & "WHERE (((con_meses.id)=" & xMes & "))", xCon
        
        fgx.TextMatrix(0, contador + 3) = Rst("descripcion") & " " & xAño
        fgx.ColWidth(contador + 3) = 1250
        xMes = xMes + 1
        If xMes > 12 Then xMes = 1: xAño = xAño + 1
        contador = contador + 1
        If contador > indicador Then bandera = False
    Wend
    fgx.TextMatrix(0, contador + 3) = "Programado"
    fgx.ColWidth(contador + 3) = 1200
    fgx.TextMatrix(0, contador + 4) = "Stock Ini."
    fgx.Cols = fgx.Cols + 2
    fgx.TextMatrix(0, contador + 5) = "Producido"
    fgx.TextMatrix(0, contador + 6) = "Total"
    fgx.ColWidth(contador + 6) = 1200
    fgx.TextMatrix(0, contador + 7) = "Diferencia"
    Set Rst = Nothing
End Sub

'*******************************************************************************************
'Nombre Sub o Function: MuestraSegundoTab
'Tipo:                  PROCEDIMIENTO
'Descripcion:           MUESTRA INFORMACION DETALLADA DEL REGISTRO SELECCIONADO
'Hecho por:             ENRIQUE POLLONGO SIERRA
'Modificado:            Por Jose Chacon:
'                           -Remodelamiento Total del Procedimiento para implementacion de
'                           Tablas Cruzadas
'*******************************************************************************************
Sub MuestraSegundoTab()
    Dim Rst As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    Dim Rst2Aux As New ADODB.Recordset
    Dim A, B, xCol As Integer
    Dim Total As Double
    
    Dim idMesIni As Integer
    Dim idMesFin As Integer
    Dim idAñoIni As Integer
    Dim idAñoFin As Integer
    
    Dim xMes As Integer
    Dim xAño As Integer
    Dim bandera As Boolean
    Dim contador As Integer
    
    Dim xMesAux As String
    
    Fg1.Rows = 1
    Fg1.Cols = 3
    
    Fg2.Rows = 1
    Fg2.Cols = 3
    
    TxtDesc.Text = RstPlaPro("descripcion")
    TxtFchIni.Valor = Format(RstPlaPro("fchini"), "dd/mm/yyyy")
    TxtFchFin.Valor = Format(RstPlaPro("fchfin"), "dd/mm/yyyy")
    
    idMesIni = Format(TxtFchIni.Valor, "m")
    'idMesFin = CInt(Mid(TxtFchFin.Valor, 4, 2))
    idAñoIni = Format(TxtFchIni.Valor, "yyyy")
    'idAñoFin = CInt(Mid(TxtFchFin.Valor, 7, 4))
    
    xMes = idMesIni
    xAño = idAñoIni
    
    Dim indicador As Integer
    indicador = calcularIndicador(CDate(RstPlaPro("fchini")), CDate(RstPlaPro("fchfin")))
    
    Fg1.Cols = Fg1.Cols + indicador + 4
    
    RellenarMeses Fg1, CDate(TxtFchIni.Valor), CDate(TxtFchFin.Valor)
    TabOne2.CurrTab = 0
    
    'MOSTRAMOS LOS PRODUCTOS FINALES
    RST_Busq Rst, "TRANSFORM First(ges_plaproddet.cantidad) AS PrimeroDecantidad " _
        + vbCr + "SELECT ges_plaproddet.codpro, alm_inventario.codpro AS codigo, alm_inventario.descripcion, mae_unidades.abrev, ges_plaproddet.idpv, alm_inventario.stckini " _
        + vbCr + "FROM (ges_plaproddet INNER JOIN alm_inventario ON ges_plaproddet.codpro = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        + vbCr + "Where (((ges_plaproddet.idpv) = " & RstPlaPro("id") & ") And ((ges_plaproddet.idmes) <> 13)) " _
        + vbCr + "GROUP BY ges_plaproddet.codpro, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ges_plaproddet.idpv, alm_inventario.stckini " _
        + vbCr + "PIVOT ges_plaproddet.idmes", xCon

    Fg1.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Total = 0
            Fg1.TextMatrix(A, 0) = NulosC(Rst("codpro"))
            Fg1.TextMatrix(A, 1) = NulosC(Rst("descripcion"))
            Fg1.TextMatrix(A, 2) = NulosC(Rst("codigo"))
            Fg1.TextMatrix(A, 3) = NulosC(Rst("abrev"))
            
            xMesAux = xMes
            For B = 1 To indicador
                Fg1.TextMatrix(A, B + 3) = Format(NulosN(Rst("" & xMesAux & "")), FORMAT_MONTO)
                Total = Total + NulosN(Rst("" & xMesAux & ""))
                xMesAux = xMesAux + 1
                If xMesAux > 12 Then xMesAux = 1
            Next B
            
            Fg1.TextMatrix(A, B + 3) = Format(Total, FORMAT_MONTO)
            Rst.MoveNext
        Next A
        Set Rst = Nothing
    End If
    
    'MOSTRAMOS LOS PRODUCTOS INTERMEDIOS
    
    Fg2.Cols = Fg2.Cols + indicador + 4
    RellenarMeses Fg2, TxtFchIni.Valor, TxtFchFin.Valor

    RST_Busq Rst, "TRANSFORM First(ges_plaproddet2.cantidad) AS PrimeroDecantidad " _
            & "SELECT ges_plaproddet2.codpro, alm_inventario.codpro AS codigo, alm_inventario.descripcion, mae_unidades.abrev, ges_plaproddet2.idpv, alm_inventario.stckini " _
            & "FROM (ges_plaproddet2 INNER JOIN alm_inventario ON ges_plaproddet2.codpro = alm_inventario.id) INNER JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
            & "Where (((ges_plaproddet2.idpv) = " & RstPlaPro("id") & ") And ((ges_plaproddet2.idmes) <> 13)) " _
            & "GROUP BY ges_plaproddet2.codpro, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, ges_plaproddet2.idpv, alm_inventario.stckini " _
            & "PIVOT ges_plaproddet2.idmes", xCon

    Fg2.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Total = 0
            Fg2.TextMatrix(A, 0) = NulosC(Rst("codpro"))
            Fg2.TextMatrix(A, 1) = NulosC(Rst("descripcion"))
            Fg2.TextMatrix(A, 2) = NulosC(Rst("codigo"))
            Fg2.TextMatrix(A, 3) = NulosC(Rst("abrev"))
            
            xMesAux = xMes
            For B = 1 To indicador
                Fg2.TextMatrix(A, B + 3) = Format(NulosN(Rst("" & xMesAux & "")), FORMAT_MONTO)
                Total = Total + NulosN(Rst("" & xMesAux & ""))
                xMesAux = xMesAux + 1
                If xMesAux > 12 Then xMesAux = 1
                
            Next B
            Fg2.TextMatrix(A, B + 3) = Format(Total, FORMAT_MONTO)
            Rst.MoveNext
        Next A
    End If
    
    LblNumReg.Caption = Fg1.Rows - 1
    MostrarSaldo
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 2 Then
        If ButtonMenu.Index = 1 Then Modificar
        If ButtonMenu.Index = 2 Then CambiarEstado True
    End If
    
    If ButtonMenu.Parent.Index = 3 Then
        If ButtonMenu.Index = 1 Then Eliminar
        If ButtonMenu.Index = 2 Then CambiarEstado False
    End If
End Sub

Private Sub TxtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : ExportarExcel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA A MS EXCEL LOS DATOS DEL CONTROL Fg1
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub ExportarExcel()
    Dim A As Integer
    Dim B As Integer
    Dim xFilas As Integer
    Dim xCad As String
    Dim objExcel As Object
    
    Set objExcel = CreateObject("Excel.Application")
    'Dim objExcel As New Excel.Application
    
    objExcel.Visible = True
    'determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1
    
    objExcel.WindowState = 2
    objExcel.Workbooks.Add
   
    With objExcel.ActiveSheet
        .Cells(1, 2) = NomEmp
        .Cells(1, 10) = Date
        .Cells(2, 2) = "Nº R.U.C. : " + NumRUC
        If TabOne2.CurrTab = 0 Then
            .Cells(3, 2) = "Plan de Produccion de Productos Terminados"
        Else
            .Cells(3, 2) = "Plan de Produccion de Productos Intermedios"
        End If
        
        xFilas = 5
        If TabOne2.CurrTab = 0 Then
            For A = 0 To Fg1.Rows - 1
                For B = 1 To Fg1.Cols - 1
                    If A = 0 Then
                        .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(A, B)
                    Else
                        If B <= 3 Then
                            .Cells(xFilas, B + 1) = "'" + Fg1.TextMatrix(A, B)
                        Else
                            .Cells(xFilas, B + 1) = NulosN(Fg1.TextMatrix(A, B))
                        End If
                    End If
                Next B
                xFilas = xFilas + 1
            Next A
        Else
            For A = 0 To Fg2.Rows - 1
                For B = 1 To Fg2.Cols - 1
                    If A = 0 Then
                        .Cells(xFilas, B + 1) = "'" + Fg2.TextMatrix(A, B)
                    Else
                        If B <= 3 Then
                            .Cells(xFilas, B + 1) = "'" + Fg2.TextMatrix(A, B)
                        Else
                            .Cells(xFilas, B + 1) = NulosN(Fg2.TextMatrix(A, B))
                        End If
                    End If
                Next B
                xFilas = xFilas + 1
            Next A
        End If
    End With
    
    MsgBox "El proceso de exportacion termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Registro de Compras y Ventas"
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
End Sub
