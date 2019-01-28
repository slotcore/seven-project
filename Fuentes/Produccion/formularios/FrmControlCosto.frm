VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmControlCosto 
   Caption         =   "Producción - Ingreso de Costos"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7245
      Left            =   15
      TabIndex        =   0
      Top             =   360
      Width           =   11895
      _cx             =   20981
      _cy             =   12779
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
         Caption         =   "Detalle de la Cuenta"
         Height          =   6825
         Left            =   12540
         TabIndex        =   6
         Top             =   375
         Width           =   11805
         Begin VB.CommandButton CmdRecalcular 
            Caption         =   "Recalcular"
            Enabled         =   0   'False
            Height          =   330
            Index           =   4
            Left            =   8100
            TabIndex        =   36
            TabStop         =   0   'False
            ToolTipText     =   "Buscar Tarea o Receta"
            Top             =   615
            Width           =   1770
         End
         Begin VB.TextBox txt 
            BackColor       =   &H0080FF80&
            Height          =   315
            Index           =   0
            Left            =   9435
            TabIndex        =   10
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   75
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cb 
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   4605
            Picture         =   "FrmControlCosto.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   675
            Width           =   225
         End
         Begin VB.CommandButton cb 
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   4605
            Picture         =   "FrmControlCosto.frx":0132
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   345
            Width           =   225
         End
         Begin VB.Frame Frame4 
            Caption         =   "( Periodo )"
            Height          =   615
            Left            =   10020
            TabIndex        =   7
            Top             =   -15
            Width           =   1740
            Begin VB.Label lblperiodo 
               Alignment       =   2  'Center
               Caption         =   "lblperiodo"
               BeginProperty Font 
                  Name            =   "Courier"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   8
               Top             =   315
               Width           =   1605
            End
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
            Height          =   300
            Index           =   0
            Left            =   855
            TabIndex        =   11
            Top             =   315
            Width           =   1230
            _ExtentX        =   2170
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
            Valor           =   "21/11/2007"
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   1
            Left            =   3645
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   13
            Text            =   "txt_cb(1)"
            Top             =   645
            Width           =   1215
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   0
            Left            =   3645
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   14
            Text            =   "txt_cb(0)"
            Top             =   315
            Width           =   1215
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
            Height          =   300
            Index           =   1
            Left            =   855
            TabIndex        =   23
            Top             =   675
            Width           =   1230
            _ExtentX        =   2170
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
            Valor           =   "21/11/2007"
         End
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   5745
            Left            =   -15
            TabIndex        =   25
            Top             =   1020
            Width           =   11850
            _cx             =   20902
            _cy             =   10134
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
            FrontTabColor   =   -2147483644
            BackTabColor    =   -2147483633
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483630
            Caption         =   "     Detalle     |     Resumen     "
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
            Begin VB.Frame Frame3 
               BorderStyle     =   0  'None
               Caption         =   "Frame3"
               Height          =   5325
               Left            =   45
               TabIndex        =   28
               Top             =   45
               Width           =   11760
               Begin VB.Frame Frame5 
                  Caption         =   "Seleccionar Concepto de Ingreso >> Planillas"
                  Height          =   765
                  Left            =   4410
                  TabIndex        =   37
                  Top             =   4500
                  Width           =   5475
                  Begin VB.CommandButton cb 
                     Enabled         =   0   'False
                     Height          =   240
                     Index           =   2
                     Left            =   1875
                     Picture         =   "FrmControlCosto.frx":0264
                     Style           =   1  'Graphical
                     TabIndex        =   38
                     Top             =   345
                     Width           =   225
                  End
                  Begin VB.TextBox txt_cb 
                     Height          =   300
                     Index           =   2
                     Left            =   915
                     Locked          =   -1  'True
                     MaxLength       =   12
                     TabIndex        =   39
                     Text            =   "txt_cb(2)"
                     Top             =   315
                     Width           =   1215
                  End
                  Begin VB.Label lbl_cod 
                     BackColor       =   &H000000FF&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "lbl_cod(2)"
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
                     Index           =   2
                     Left            =   3120
                     TabIndex        =   41
                     Top             =   300
                     Visible         =   0   'False
                     Width           =   1185
                  End
                  Begin VB.Label lbl_cb_capt 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Concepto"
                     Height          =   195
                     Index           =   2
                     Left            =   90
                     TabIndex        =   40
                     Top             =   420
                     Width           =   690
                  End
                  Begin VB.Label lbl_cb 
                     BackStyle       =   0  'Transparent
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "lbl_cb(2)"
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
                     Index           =   2
                     Left            =   2115
                     TabIndex        =   42
                     Top             =   315
                     Width           =   3180
                  End
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "Eliminar Otros Ingresos"
                  Height          =   330
                  Index           =   2
                  Left            =   2250
                  TabIndex        =   32
                  TabStop         =   0   'False
                  ToolTipText     =   "Buscar Tareas Realizadas"
                  Top             =   4590
                  Width           =   2115
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "&Agregar Otros Ingresos"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   0
                  Left            =   0
                  TabIndex        =   31
                  ToolTipText     =   "Agregar "
                  Top             =   4590
                  Width           =   2115
               End
               Begin VSFlex7Ctl.VSFlexGrid Fg2 
                  Height          =   4455
                  Left            =   30
                  TabIndex        =   29
                  Top             =   30
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   7858
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
                  ForeColorSel    =   16777215
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   1
                  Cols            =   2
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmControlCosto.frx":0396
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
               Begin VB.CommandButton Cmd 
                  Caption         =   "Agregar Descuento"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   1
                  Left            =   0
                  TabIndex        =   30
                  TabStop         =   0   'False
                  ToolTipText     =   "Eliminar "
                  Top             =   4950
                  Width           =   2115
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "Eliminar Descuento"
                  Height          =   330
                  Index           =   3
                  Left            =   2250
                  TabIndex        =   33
                  TabStop         =   0   'False
                  ToolTipText     =   "Exportar MSExcel"
                  Top             =   4950
                  Width           =   2115
               End
            End
            Begin VB.Frame Frame6 
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   5325
               Left            =   -12405
               TabIndex        =   26
               Top             =   45
               Width           =   11760
               Begin VSFlex7Ctl.VSFlexGrid Fg1 
                  Height          =   5265
                  Left            =   15
                  TabIndex        =   27
                  Top             =   15
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   9287
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
                  ForeColorSel    =   16777215
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
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   1
                  Cols            =   2
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmControlCosto.frx":05A7
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
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fch Final"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   24
            Top             =   750
            Width           =   645
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Index           =   0
            Left            =   8880
            TabIndex        =   17
            Top             =   195
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supervizado Por"
            Height          =   195
            Index           =   1
            Left            =   2430
            TabIndex        =   21
            Top             =   750
            Width           =   1170
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fch Inicio"
            Height          =   195
            Index           =   8
            Left            =   60
            TabIndex        =   19
            Top             =   420
            Width           =   690
         End
         Begin VB.Label lbl_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cod(1)"
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
            Index           =   1
            Left            =   5850
            TabIndex        =   18
            Top             =   630
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label lbl_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cod(0)"
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
            Index           =   0
            Left            =   5850
            TabIndex        =   16
            Top             =   315
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Programado Por"
            Height          =   195
            Index           =   0
            Left            =   2430
            TabIndex        =   15
            Top             =   420
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Detalle del Ingreso de Costos del Personal"
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
            Left            =   75
            TabIndex        =   20
            Top             =   15
            Width           =   11610
         End
         Begin VB.Label lbl_cb 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb(0)"
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
            Index           =   0
            Left            =   4845
            TabIndex        =   34
            Top             =   315
            Width           =   3180
         End
         Begin VB.Label lbl_cb 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb(1)"
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
            Index           =   1
            Left            =   4845
            TabIndex        =   35
            Top             =   645
            Width           =   3180
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6825
         Left            =   45
         TabIndex        =   1
         Top             =   375
         Width           =   11805
         Begin TrueOleDBGrid70.TDBGrid Dg3 
            Height          =   6465
            Left            =   45
            TabIndex        =   2
            Top             =   345
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11404
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
            Columns(1).Caption=   "Fch Trabajo"
            Columns(1).DataField=   "fchtra"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Area"
            Columns(2).DataField=   "area"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Responsable"
            Columns(3).DataField=   "encargado"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2117"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2037"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=4524"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=4445"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=7673"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=7594"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14,.alignment=2"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Named:id=33:Normal"
            _StyleDefs(53)  =   ":id=33,.parent=0"
            _StyleDefs(54)  =   "Named:id=34:Heading"
            _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(56)  =   ":id=34,.wraptext=-1"
            _StyleDefs(57)  =   "Named:id=35:Footing"
            _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(59)  =   "Named:id=36:Selected"
            _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(61)  =   "Named:id=37:Caption"
            _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(63)  =   "Named:id=38:HighlightRow"
            _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(65)  =   "Named:id=39:EvenRow"
            _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(67)  =   "Named:id=40:OddRow"
            _StyleDefs(68)  =   ":id=40,.parent=33"
            _StyleDefs(69)  =   "Named:id=41:RecordSelector"
            _StyleDefs(70)  =   ":id=41,.parent=34"
            _StyleDefs(71)  =   "Named:id=42:FilterBar"
            _StyleDefs(72)  =   ":id=42,.parent=33"
         End
         Begin VB.Label LblMes 
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
            ForeColor       =   &H00000080&
            Height          =   300
            Left            =   8835
            TabIndex        =   5
            Top             =   30
            Width           =   1275
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Costos del Personal"
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
            Left            =   105
            TabIndex        =   4
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label lblperiodo 
            AutoSize        =   -1  'True
            Caption         =   "lblperiodo"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   0
            Left            =   9705
            TabIndex        =   3
            Top             =   30
            Width           =   1095
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   1005
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
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Listado"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Registro"
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
         Left            =   6600
         Top             =   0
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
               Picture         =   "FrmControlCosto.frx":07B8
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlCosto.frx":0CFC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlCosto.frx":108E
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlCosto.frx":1212
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlCosto.frx":1666
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlCosto.frx":177E
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlCosto.frx":1CC2
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlCosto.frx":2206
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlCosto.frx":231A
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlCosto.frx":242E
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlCosto.frx":2882
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlCosto.frx":29EE
               Key             =   "IMG11"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmControlCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim QueHace As Integer
'Dim Agregando As Boolean
'Dim SeEjecuto As Boolean
'Dim RstFrm As New ADODB.Recordset
''----
'Dim mMesActivo  As Integer
'
'Dim fOrdenLista As Boolean  '--especfica el orden de la lista de la consulta
'Dim mRowAdd As Double       '--identificador unico por fila cuando se agrege una tarea
'Dim mRowAddTara As Double   '--identificador unico por fila cuando se agrege una tarea
'
'Dim mIdRegistro&            '--identificador del registro
'Dim sPesoTara As Double     '--indica la perdida de peso segun unidad
'
'Public RstGrDet As New ADODB.Recordset  '--
'Public RstGrDetTara As New ADODB.Recordset  '--
'
'Private fActivarAutomaticoCantidad As Boolean '--permitira controlar la distribucion de las cantidades pro grupo
'Private fActivarAutomaticoHora As Boolean '--permitira controlar la distribucion de las horas pro grupo
'
'Private Sub chkOpcion_Click(Index As Integer)
'    Select Case Index
'        Case 2 '--tipo
'            chkOpcion(3).Value = 0
'
'        Case 3 '--
'            If chkOpcion(3).Value = 1 Then
'                If chkOpcion(2).Value = 0 Then
'                    chkOpcion(3).Value = 0
'                End If
'            End If
'
'        Case 4 '--tarea
'            chkOpcion(5).Value = 0
'
'        Case 5 '--producto
'            If chkOpcion(5).Value = 1 Then
'                If chkOpcion(4).Value = 0 Then
'                    chkOpcion(5).Value = 0
'                End If
'            End If
'
'    End Select
'End Sub
'
'Private Sub Cmd_Click(Index As Integer)
'
'    Select Case Index
'        Case 0 '--agregar
'            pRegistroAdd
'        Case 1 '--eliminar
'            pRegistroDel
'        Case 2 '--mostrar la consulta de tareas en las recetas
'            pHabilitarBotonEditor 2, True
'        Case 3 '--cuadro de opciones
'            pHabilitarBotonEditor 3, True
'    End Select
'End Sub
'
'Private Sub CmdTarea_Click(Index As Integer)
'    pHabilitarBotonEditor 2, False
'End Sub
'
'Private Sub CmdUtil_Click(Index As Integer)
'    Select Case Index
'        Case 0 '--buscar registros
'            pBuscarVSFlexGrid
'        Case 1 '--exportar msexcel vsflexgrid
'            pExportarVSFlexGrid
'    End Select
'End Sub
'
'
'Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 2 Then
'        If QueHace <> 3 Then
''            If Index = 0 Then PopupMenu Menu3
''            If Index = 1 Then PopupMenu menu2
'        End If
'    End If
'
'End Sub
'
'Private Sub Dg3_DblClick()
'    TabOne1.CurrTab = 1
'End Sub
'
'Private Sub Dg3_HeadClick(ByVal ColIndex As Integer)
'    On Error Resume Next
'    Dim nOrden As String
'    If fOrdenLista = False Then nOrden = "ASC"
'    If fOrdenLista = True Then nOrden = "DESC"
'    RstFrm.Sort = CStr(Dg3.Columns(ColIndex).DataField) & " " & nOrden
'    fOrdenLista = Not fOrdenLista
'    Err.Clear
'End Sub
'
'Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
'    On Error GoTo error
'    If Agregando = True Then Exit Sub
'    If Row = 0 Then Exit Sub
'
'    fActivarAutomaticoHora = False
'    fActivarAutomaticoCantidad = False
'
'    Select Case Col
'        Case 2 '--tipo persona
'            '--actualizando el codigo
'            If NulosN(Fg1.TextMatrix(Row, 11)) <> NulosN(Fg1.TextMatrix(Row, 2)) And NulosN(Fg1.TextMatrix(Row, 11)) <> 0 Then
'
'                '--eliminar los registros del grupo anterior
'                If NulosN(Fg1.TextMatrix(Row, 12)) <> 0 Then RstRegistroEliminar RstGrDet, "codigo", NulosN(Fg1.TextMatrix(Row, 16)), True
'                Fg1.TextMatrix(Row, 3) = "": '--personal / grupo
'                Fg1.TextMatrix(Row, 12) = "" '--idref personal / grupo
'
'            End If
'            Fg1.TextMatrix(Row, 11) = NulosN(Fg1.TextMatrix(Row, 2))
'
'        Case 3 '--personal(individual/grupal)
'            '--limpiar CodRef
'            If NulosC(Fg1.TextMatrix(Row, Col)) = "" Then Fg1.TextMatrix(Row, 12) = ""
'        Case 4 '--tarea
'            '--limpiar CodTarea
'            If NulosC(Fg1.TextMatrix(Row, Col)) = "" Then
'                Fg1.TextMatrix(Row, 14) = "" '--idtarea
'
'                Fg1.TextMatrix(Row, 5) = "" '--producto
'                Fg1.TextMatrix(Row, 13) = "" '--idrec
'            End If
'
'        Case 5 '--producto
'            '--limpiar CodReceta
'            If NulosC(Fg1.TextMatrix(Row, Col)) = "" Then Fg1.TextMatrix(Row, 13) = ""
'
'        Case 6, 7
'            If Fg1.TextMatrix(Row, Col) = "  :  " Then
'                Fg1.TextMatrix(Row, Col) = "":  GoTo Continuar1:
'            End If
'            If IsDate(Fg1.TextMatrix(Row, Col)) = False Then
'                MsgBox "El valor ingresado no es una Hora correcta", vbCritical, xTitulo
'                Fg1.TextMatrix(Row, Col) = ""
'            Else
''--se valida para que se ingrese las horas correctas
'''                If IsDate(Fg1.TextMatrix(Row, 6)) = True And IsDate(Fg1.TextMatrix(Row, 7)) = True Then '--HORA INICIO
'''                    If CDate(Fg1.TextMatrix(Row, 6)) >= CDate(Fg1.TextMatrix(Row, 7)) Then
'''                        MsgBox "La hora " + IIf(Col = 6, "Inicial debe ser menor ", "Final debe ser mayor") + " a la hora " + IIf(Col = 6, "Final", "Inicial"), vbExclamation, xTitulo
'''                        Fg1.TextMatrix(Row, Col) = "":  Exit Sub
'''                   End If
'''                End If
'                Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), FORMAT_HORA_SIN_SEGUNDO)
'
'                '--------------
'
'Continuar1:
'                '--actualizar si es grupo
'                If NulosN(Fg1.TextMatrix(Row, 2)) = 2 Then
'
'                    If BuscarFrm("FrmControlTareaGr", True) = True Then
'                        If FrmControlTareaGr.fDesactivarAuto = False Then fActivarAutomaticoHora = True
'                    Else
'                        fActivarAutomaticoHora = True
'                    End If
'
'                    Fg1_RowColChange
'                    fActivarAutomaticoHora = False
'                End If
'
'
'            End If
'
'        Case 8 '--cantidad
'            If NulosC(Fg1.TextMatrix(Row, Col)) = "" Then
'                    '--actualizar si es grupo
'                    fActivarAutomaticoCantidad = True
'                    If NulosN(Fg1.TextMatrix(Row, 2)) = 2 Then Fg1_RowColChange
'                    fActivarAutomaticoCantidad = False
'            Else
'                If IsNumeric(Fg1.TextMatrix(Row, Col)) = False Then
'                    MsgBox "El valor ingresado no es numérico", vbCritical, xTitulo
'                    Fg1.TextMatrix(Row, Col) = ""
'                Else
'                    Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), FORMAT_CANTIDAD)
'                    '--actualizar si es grupo
'                    If NulosN(Fg1.TextMatrix(Row, 2)) = 2 Then
'
'                        If BuscarFrm("FrmControlTareaGr", True) = True Then
'                            If FrmControlTareaGr.fDesactivarAuto = False Then
'                                If NulosN(Fg1.TextMatrix(Row, Col)) <> NulosN(Fg1.TextMatrix(Row, 17)) Then fActivarAutomaticoCantidad = True
'                            End If
'                        Else
'                            If NulosN(Fg1.TextMatrix(Row, Col)) <> NulosN(Fg1.TextMatrix(Row, 17)) Then fActivarAutomaticoCantidad = True
'                        End If
'
'                        Fg1_RowColChange
'
'                    End If
'                End If
'            End If
'
'        Case 9 '--unidad de medida
'            '--limpiar CodTarea
'            If NulosC(Fg1.TextMatrix(Row, Col)) = "" Then Fg1.TextMatrix(Row, 15) = ""
'
'    End Select
'    Exit Sub
'error:
'    SHOW_ERROR Me.Name, "Fg1_CellChanged ( " & Row & "," & Col & ")"
'End Sub
'
'Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'
'    If Col <> 3 And Col <> 4 And Col <> 5 And Col <> 8 And Col <> 9 Then Exit Sub
'    If QueHace = 3 Then Exit Sub
'    Dim RstTmp As New ADODB.Recordset
'    Dim nSQL As String
'    Dim nSQLTmp As String
'    Dim nSQLNotId As String
'    Dim nTitulo As String
'
'    On Error GoTo error
'
'    Select Case Col
'        Case 2 '--tipo persona
'
'        Case 3 '--personal / nº grupo
'            If NulosN(Fg1.TextMatrix(Row, 2)) = 0 Then '--indivual
'                MsgBox "Seleccione el Tipo Individual o Grupal", vbExclamation, xTitulo
'                Fg1.Col = 2
'                Fg1.SetFocus
'                Exit Sub
'            End If
'            If NulosN(Fg1.TextMatrix(Row, 2)) = 1 Then '--individual
'
'                If NulosC(Fg1.TextMatrix(Row, Col)) <> "" Then
'                    nSQLTmp = " AND UCASE([pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom]) LIKE '%" & UCase(NulosC(Fg1.TextMatrix(Row, Col))) & "%'"
'                End If
'
'                ReDim xCampos(2, 4) As String
'                xCampos(0, 0) = "Apellidos y Nombres":  xCampos(0, 1) = "nombre":      xCampos(0, 2) = "4500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
'                xCampos(1, 0) = "Fch. Nac":             xCampos(1, 1) = "fchnac":      xCampos(1, 2) = "1000":     xCampos(1, 3) = "C":    xCampos(1, 4) = "D"
'
'                nSQL = "SELECT pla_empleados.id AS idemp, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, pla_empleados.fchnac " _
'                    + vbCr + " FROM (pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper " _
'                    + vbCr + " Where (((pro_empdet.idfun) = 6 )) " & nSQLTmp _
'                    + vbCr + " ORDER BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom]; "
'
'                nTitulo = "Buscando Personal"
'
'            Else '--grupo
'                ReDim xCampos(3, 4) As String
'                xCampos(0, 0) = "Nº Grupo":         xCampos(0, 1) = "nombre":     xCampos(0, 2) = "900":   xCampos(0, 3) = "C":
'                xCampos(1, 0) = "Responsable":      xCampos(1, 1) = "encargado":  xCampos(1, 2) = "3500":  xCampos(1, 3) = "C":
'                xCampos(2, 0) = "Nº Integrantes":   xCampos(2, 1) = "totper":     xCampos(2, 2) = "1400":  xCampos(2, 3) = "N":
'
'                If NulosN(Fg1.TextMatrix(Row, 12)) <> 0 Then nSQLNotId = " WHERE pro_grupo.id <> " & NulosN(Fg1.TextMatrix(Row, 12))
'
'                nSQL = "SELECT pro_grupo.id as idgru, pro_grupo.num as nombre , 'GRUPO Nº' &  pro_grupo.num as referencia, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS encargado, Count(pro_grupodet.idgrupo) AS totper " _
'                    + vbCr + " FROM (pro_grupo LEFT JOIN (pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) ON pro_grupo.idres = pro_emp.id) INNER JOIN pro_grupodet ON pro_grupo.id = pro_grupodet.idgrupo " _
'                    + vbCr + nSQLNotId & " GROUP BY pro_grupo.id, pro_grupo.num, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom];"
'
'                nTitulo = "Buscando Grupos"
'
'            End If
'
'        Case 4 '--de la tarea
'
'                ReDim xCampos(4, 4) As String
'                xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":     xCampos(0, 2) = "4500":    xCampos(0, 3) = "C"
'                xCampos(1, 0) = "Abrev":        xCampos(1, 1) = "nomcorto":   xCampos(1, 2) = "2300":    xCampos(1, 3) = "C"
'                xCampos(2, 0) = "Diverso":      xCampos(2, 1) = "diverso":    xCampos(2, 2) = "700":     xCampos(2, 3) = "C"
'                xCampos(3, 0) = "Id":           xCampos(3, 1) = "id":         xCampos(3, 2) = "600":     xCampos(3, 3) = "N"
'
'                If NulosC(Fg1.TextMatrix(Row, Col)) <> "" Then
'                    nSQLTmp = " AND (UCASE(pro_tareas.descripcion) LIKE '%" & UCase(NulosC(Fg1.TextMatrix(Row, Col))) & "%' OR UCASE(pro_tareas.abrev) LIKE '%" & UCase(NulosC(Fg1.TextMatrix(Row, Col))) & "%' ) "
'                End If
'
'                '--si hay area seleccionada o que el filtro de areas seleccioanadas este desacivadas
'                If NulosN(lbl_cod(0).Caption) <> 0 And chkOpcion(0).Value = 0 Then
'                    nSQL = "SELECT pro_tareas.id, pro_tareas.codigo, pro_tareas.descripcion AS nombre,pro_tareas.abrev AS nomcorto, mae_unidades.id AS idunimed, mae_unidades.abrev, IIf([pro_tareas].[diverso]=-1,'Si','No') AS diverso " _
'                        + vbCr + " FROM (mae_unidades RIGHT JOIN (pro_tareas LEFT JOIN pro_areadet ON pro_tareas.id = pro_areadet.idtar) ON mae_unidades.id = pro_tareas.idunimed) LEFT JOIN pro_area ON pro_areadet.idar = pro_area.id " _
'                        + vbCr + " WHERE pro_areadet.activo = -1 And pro_area.idarea = " & NulosN(lbl_cod(0).Caption)
'
'                Else '--no hay area seleccionada
'                    nSQLTmp = Replace(nSQLTmp, "AND", "WHERE")
'                    nSQL = "SELECT pro_tareas.id, pro_tareas.codigo, pro_tareas.descripcion AS nombre, pro_tareas.abrev AS nomcorto, mae_unidades.id AS idunimed, mae_unidades.abrev, IIf([pro_tareas].[diverso]=-1,'Si','No') AS diverso " _
'                        + vbCr + " FROM mae_unidades RIGHT JOIN pro_tareas ON mae_unidades.id = pro_tareas.idunimed  "
'
'                End If
'                nSQL = nSQL & nSQLTmp
'
'                nTitulo = "Buscando Tareas"
'
'
'        Case 5 '--del producto
'
''            If IsDate(TxtFecha(0).Valor) = False Then
''                MsgBox "Falta especificar la Fecha de Trabajo", vbExclamation, xTitulo
''                TxtFecha(0).SetFocus
''                Exit Sub
''            End If
'
'            If NulosN(Fg1.TextMatrix(Row, 14)) = 0 Then '--id tarea
'                MsgBox "Falta Especificar la Tarea", vbExclamation, xTitulo
'                Fg1.Col = 4
'                Fg1.SetFocus
'                Exit Sub
'            End If
'
'            ReDim xCampos(3, 4) As String
'            xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "codpro":   xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
'            xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "nombre":   xCampos(1, 2) = "5000":    xCampos(1, 3) = "C"
'            xCampos(2, 0) = "CodReceta":    xCampos(2, 1) = "codrec":   xCampos(2, 2) = "1200":    xCampos(2, 3) = "C"
'
'
'            If NulosC(Fg1.TextMatrix(Row, Col)) <> "" Then
'                nSQLTmp = " AND UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(Fg1.TextMatrix(Row, Col))) & "%'"
'            End If
'
'            nSQL = "SELECT DISTINCT alm_inventario.codpro, alm_inventario.descripcion as nombre, pro_receta.codrec, pro_receta.iditem, pro_receta.id AS idrec " _
'                + vbCr + " FROM alm_inventario INNER JOIN pro_receta ON alm_inventario.id = pro_receta.iditem " _
'                + vbCr + " WHERE pro_receta.id IN (SELECT pro_recetatar.idrec FROM pro_recetatar WHERE pro_recetatar.idtar= " & NulosN(Fg1.TextMatrix(Row, 14)) & ") " & nSQLTmp _
'                + vbCr + " ORDER BY alm_inventario.descripcion; "
'
'            nTitulo = "Buscando Productos"
'
'        Case 9 '--unidad de medida
'            ReDim xCampos(2, 4) As String
'            xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":   xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
'            xCampos(1, 0) = "Abrev":        xCampos(1, 1) = "abrev":    xCampos(1, 2) = "800":    xCampos(1, 3) = "C"
'            xCampos(2, 0) = "Id":           xCampos(2, 1) = "id":       xCampos(2, 2) = "600":    xCampos(2, 3) = "N"
'
'            nSQL = "SELECT mae_unidades.id, mae_unidades.descripcion as nombre, mae_unidades.abrev FROM mae_unidades;"
'        Case 8
'            pHabilitarBotonEditor 1, True
'            Exit Sub
'        Case Else
'            Exit Sub
'    End Select
'
'    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio, ""
'
'    If RstTmp.State = 0 Then GoTo SALIR
'    If RstTmp.EOF = True Or RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR
'    Agregando = True
'
'    If Col = 2 Then '--tipo persona
'        'Fg1.TextMatrix(Row, Col) = NulosC(RstTmp.Fields("nombre"))
'    ElseIf Col = 3 Then '--persona / grupo
'
'        '-------------------------------------------------------------------
'        If NulosN(Fg1.TextMatrix(Row, 2)) = 2 Then '--solo grupos
'            '--si el grupo inicial es diferente al grupo actual
'            If NulosN(Fg1.TextMatrix(Row, 12)) <> NulosN(RstTmp.Fields("idgru")) Then
'                '--eliminar los registros del grupo anterior
'                RstGrDet.Filter = ""
'                If NulosN(Fg1.TextMatrix(Row, 12)) <> 0 Then RstRegistroEliminar RstGrDet, "codigo", NulosN(Fg1.TextMatrix(Row, 16)), True
'                '--cargar los datos del grupo
'                pCargarDatosRstTemp 0, NulosN(RstTmp.Fields("idgru")), NulosN(Fg1.TextMatrix(Row, 16)), False
'                Fg1.TextMatrix(Row, 12) = NulosN(RstTmp.Fields("idgru"))
'                '--mostrar los datos en la ventana
'                Agregando = False
'                fActivarAutomaticoCantidad = True
'                Fg1_RowColChange
'                fActivarAutomaticoCantidad = False
'                Agregando = True
'            End If
'            Fg1.TextMatrix(Row, Col) = NulosC(RstTmp.Fields("referencia"))
'
'        Else
'            Fg1.TextMatrix(Row, Col) = NulosC(RstTmp.Fields("nombre"))
'            Fg1.TextMatrix(Row, 12) = NulosN(RstTmp.Fields("idemp"))
'
'        End If
'        '-------------------------------------------------------------------
'        Fg1.Col = 4
'
'    ElseIf Col = 4 Then '--tarea
'
'        '--si la tarea es diferente => limpiar producto
'        If NulosN(Fg1.TextMatrix(Row, 13)) <> 0 And (NulosN(Fg1.TextMatrix(Row, 14)) <> NulosN(RstTmp.Fields("id"))) Then
'            Fg1.TextMatrix(Row, 5) = ""
'            Fg1.TextMatrix(Row, 13) = ""
'        End If
'
'        Fg1.TextMatrix(Row, Col) = NulosC(RstTmp.Fields("nomcorto"))
'        Fg1.TextMatrix(Row, 14) = NulosN(RstTmp.Fields("id"))
'        '--agregando la unidad por defecto
'        Fg1.TextMatrix(Row, 9) = NulosC(RstTmp.Fields("abrev"))
'        Fg1.TextMatrix(Row, 15) = NulosN(RstTmp.Fields("idunimed"))
'        Fg1.Col = 5
'
'    ElseIf Col = 5 Then '--producto
'        Fg1.TextMatrix(Row, Col) = NulosC(RstTmp.Fields("nombre"))
'        Fg1.TextMatrix(Row, 13) = NulosN(RstTmp.Fields("idrec"))
'        Fg1.Col = 6
'
'    ElseIf Col = 9 Then '--unidad de medida
'        Fg1.TextMatrix(Row, Col) = NulosC(RstTmp.Fields("abrev"))
'        Fg1.TextMatrix(Row, 15) = NulosN(RstTmp.Fields("id"))
'        Fg1.Col = 10
'
'    End If
'
'    Agregando = False
'    Set RstTmp = Nothing
'    Exit Sub
'SALIR:
'    If Col = 3 Then '--persona / grupo
'        Fg1.Col = 3
'    ElseIf Col = 4 Then '--producto
'        Fg1.Col = 4
'    ElseIf Col = 5 Then '--tarea
'        Fg1.Col = 5
'    ElseIf Col = 9 Then '--unidad de medida
'        Fg1.Col = 9
'    End If
'    Fg1.TextMatrix(Row, Col) = ""
'    Fg1.SetFocus
'    Set RstTmp = Nothing
'    Agregando = False
'    Exit Sub
'error:
'    Set RstTmp = Nothing
'    Agregando = False
'    SHOW_ERROR Me.Name, "Fg1_CellButtonClick(" & Row & "," & Col & ")"
'End Sub
'
'Private Sub Fg1_EnterCell()
'    If QueHace = 3 Then
'        Fg1.Editable = flexEDNone
'        Exit Sub
'    End If
'    If Fg1.Col <= 10 Or Fg1.Col = 18 Or Fg1.Col = 19 Then
'        Fg1.Editable = flexEDKbdMouse
'    Else
'        Fg1.Editable = flexEDNone
'    End If
'End Sub
'
'Private Sub Fg1_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF6 Then
'        If Fg1.Row < 1 Then Exit Sub
'        If NulosN(Fg1.TextMatrix(Fg1.Row, 2)) = 1 Then Exit Sub
'        If BuscarFrm("FrmControlTareaGr", True, False) = True Then
'            If FrmControlTareaGr.fg(0).Rows > 1 Then
'                FrmControlTareaGr.fg(0).Row = 1
'                FrmControlTareaGr.fg(0).Col = 5
'                FrmControlTareaGr.fg(0).SetFocus
'            Else
'                FrmControlTareaGr.cmd(0).SetFocus
'            End If
'            Exit Sub
'        End If
'    End If
'
'End Sub
'
'Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'    If Row = 0 Then
'        KeyAscii = 0
'        Exit Sub
'    End If
'    Select Case Col
'        Case 3, 4, 5
'            If validar_letras(KeyAscii) = False Then
'                If validar_numero(KeyAscii) = False Then KeyAscii = 0
'            End If
'        Case 6, 7, 8
'           If validar_numero(KeyAscii) = False Then KeyAscii = 0
'        Case 9  '--unidad
'            KeyAscii = 0
'        Case 1, 10 '--lote,comentario
'        Case Else
'            KeyAscii = 0
'    End Select
'End Sub
'
'Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
'    On Error GoTo error
'    If KeyCode = 114 Or KeyCode = vbKeyInsert Then 'F3 = Agregar Item
'        Cmd_Click 0
'    End If
'    If KeyCode = 115 Or KeyCode = vbKeyDelete Then
'        Cmd_Click 1  'F4 = Eliminar Item
'    End If
'    Exit Sub
'error:
'    SHOW_ERROR Me.Name, "Fg1_KeyUp"
'End Sub
'
'Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 2 Then
''        If QueHace = 3 Then
''            PopupMenu Menu4
''        Else
''            PopupMenu Menu1
''        End If
'    End If
'End Sub
'
'Private Sub Fg1_RowColChange()
'    If Agregando = True Then Exit Sub
'    If Fg1.Rows = 1 Then Exit Sub
''    If Fg1.Col <> 2 And Fg1.Col <> 3 And Fg1.Col <> 8 Then Exit Sub
''    Unload FrmControlTareaGr
'
'    If NulosN(Fg1.TextMatrix(Fg1.Row, 11)) = 0 Then Exit Sub '--idtipo
'    If NulosN(Fg1.TextMatrix(Fg1.Row, 12)) = 0 Then Exit Sub        '--idref
'
'    If NulosN(Fg1.TextMatrix(Fg1.Row, 2)) = 1 Then '--no es grupo
'        '--eliminar los datos registros del temporal
'        RstRegistroEliminar RstGrDet, "codigo", NulosN(Fg1.TextMatrix(Fg1.Row, 16)), True
'        Unload FrmControlTareaGr
'        Exit Sub
'    End If
'    '--mostrar en otra ventana los datos del grupo
'    FrmControlTareaGr.pRecibeLink Me.hWnd, NulosN(Fg1.TextMatrix(Fg1.Row, 16)), fActivarAutomaticoCantidad, fActivarAutomaticoHora
'    FrmControlTareaGr.Show
'    If Fg1.Enabled = True Then Fg1.SetFocus
'End Sub
'
'Private Sub Form_Activate()
'    If SeEjecuto = True Then Exit Sub
'
'    SeEjecuto = False
'    mRowAdd = -999
'    mRowAddTara = -9999
'    mMesActivo = xMes
'    pConfigurarGrilla
'    pCargarGrid
'
'    SeEjecuto = True
'    If RstFrm.State = 0 Then Exit Sub
'    If RstFrm.RecordCount = 0 Then
'        If MsgBox("No se ha registrado ningúna producción, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
'            Nuevo
'        End If
'    End If
'
'End Sub
'
'Private Sub Form_Load()
'    SeEjecuto = False
'    QueHace = 3
'
'    Dg3.Columns("fchtra").NumberFormat = FORMAT_DATE
'
'    TabOne1.CurrTab = 0
'
'    Frame1.BackColor = &H8000000F
'    Frame2.BackColor = &H8000000F
'
'    Dg3.HeadLines = 2
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    If QueHace <> 3 Then
'        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Cancel = 1
'        Exit Sub
'    Else
'
'        Set RstFrm = Nothing
'        Set RstGrDet = Nothing
'        Set RstGrDetTara = Nothing
'
'    End If
'End Sub
'
'Private Sub optTarea_Click(Index As Integer)
'    txt_cb(3).Text = ""
'    If Index = 0 Then
'        lbl_cb_capt(3).Caption = "Tarea"
'    Else
'        lbl_cb_capt(3).Caption = "Receta"
'    End If
'End Sub
'
'Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
'    Unload FrmControlTareaGr
'    If OldTab = 0 Then
'        If QueHace = 3 Then MuestraSegundoTab
'    End If
'End Sub
'
'Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'    If Button.Index = 1 Then Nuevo
'    If Button.Index = 2 Then Modificar
'    If Button.Index = 3 Then Eliminar
'
'    If Button.Index = 5 Then
'        If Grabar = True Then
'            Cancelar
'            RstFrm.Requery
'
'            If RstFrm.RecordCount <> 0 Then
'                RstFrm.MoveFirst
'                RstFrm.Find "id=" & mIdRegistro
'                If RstFrm.EOF = True Then RstFrm.MoveFirst
'            End If
'
'            Dg3.Refresh
'        End If
'    End If
'
'    If Button.Index = 6 Then Cancelar
'    If Button.Index = 8 Then Filtrar
'    If Button.Index = 9 Then RstFrm.Filter = ""
'
'    If Button.Index = 10 Then CambiarMes
'
'    If Button.Index = 11 Then Buscar
'
'    If Button.Index = 15 Then
'        Set RstFrm = Nothing
'        Unload Me
'    End If
'End Sub
'
'Sub Eliminar()
'    On Error GoTo error
'    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
'        MsgBox "No hay registros", vbExclamation, xTitulo
'        Exit Sub
'    End If
'    Dim xId&
'    xId = NulosN(RstFrm.Fields("id"))
'    TabOne1.CurrTab = 0
'    If MsgBox("¿Esta seguro de eliminar el Seguimiento de las Tareas:" & vbCr & "Area: " & NulosC(RstFrm("area")) & "?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
'        xCon.Execute "DELETe * FROM pro_controltardetgrpes WHERE idctr = " & xId & ""
'        xCon.Execute "DELETe * FROM pro_controltardetpes WHERE idctr = " & xId & ""
'        xCon.Execute "DELETe * FROM pro_controltardetgr WHERE idctr = " & xId & ""
'        xCon.Execute "DELETe * FROM pro_controltardet WHERE idctr = " & xId & ""
'        xCon.Execute "DELETe * FROM pro_controltar WHERE id = " & xId & ""
'
'        MsgBox "El Seguimiento de la tarea :" & vbCr & "Area: " & NulosC(RstFrm("area")) & vbCr & "Dia:     " & Format(RstFrm("fchtra"), "dd/mm/yy") & vbCr & "Fue eliminado con éxito", vbInformation + vbOKOnly, xTitulo
'        RstFrm.Requery
'        Dg3.Refresh
'        If RstFrm.RecordCount = 0 Then
'            If MsgBox("No hay registrado ningúna producción, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
'                Nuevo
'            End If
'        End If
'    End If
'    Exit Sub
'error:
'    SHOW_ERROR Me.Name, "Eliminar"
'End Sub
'
'Sub Cancelar()
'    QueHace = 3
'    TabOne1.TabEnabled(0) = True
'    ActivaTool
'    pHabilitarObj False
'    Label1.Caption = "Detalle del Seguimiento de Tareas"
'    Fg1.SelectionMode = flexSelectionByRow
'    Unload FrmControlTareaGr
'    TabOne1.CurrTab = 0
'    Dg3.SetFocus
'End Sub
'
'Sub Modificar()
'    If RstFrm.State = 0 Then Exit Sub
'    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
'        MsgBox "No hay Registros", vbExclamation, xTitulo
'        Exit Sub
'    End If
'
'    QueHace = 2
'
'    TabOne1.TabEnabled(0) = False
'    ActivaTool
'    pHabilitarObj True
'    If TabOne1.CurrTab = 0 Then
'        TabOne1.CurrTab = 1
'        MuestraSegundoTab
'    End If
'
'    Fg1.SelectionMode = flexSelectionFree
'
'    Label1.Caption = "Modificando Seguimiento de Tarea"
'
'    txt_cb(1).SetFocus
'End Sub
'
'Private Sub MuestraSegundoTab()
'    With RstFrm
'
'        Blanquea
'        If .State = 0 Then Exit Sub
'        If .EOF = True Or .BOF = True Or .RecordCount = 0 Then Exit Sub
'        If IsDate(.Fields("fchtra")) = True Then
'            TxtFecha(0).Valor = CDate(.Fields("fchtra"))
'        End If
'
'        txt_cb(0).Text = NulosN(RstFrm("idarea"))
'        lbl_cb(0).Caption = NulosC(RstFrm("area"))
'        lbl_cod(0).Caption = NulosN(RstFrm("idarea"))
'
'        txt_cb(1).Text = NulosC(RstFrm("numdoc"))
'        lbl_cb(1).Caption = NulosC(RstFrm("encargado"))
'        lbl_cod(1).Caption = NulosN(RstFrm("idres"))
'
'        MuestraDetalle
'
'    End With
'End Sub
'
'Private Sub MuestraDetalle()
'    Dim RstTmp As New ADODB.Recordset
'    Dim nSQL As String
''    On Error GoTo error
'    '--limpiando el rst temporal
'    Me.MousePointer = vbHourglass
'    Set RstGrDet = Nothing
'    DoEvents
'    '--------------------------------
'    nSQL = "SELECT pro_controltardet.idctr, pro_controltardet.corr, pro_controltardet.numlote, pro_controltardet.tipo, IIf([pro_controltardet].[tipo]=1,[pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom],'GRUPO Nº ' & [pro_controltardet].[idref]) AS nombres, alm_inventario.descripcion AS producto, pro_tareas.abrev AS tarea, pro_controltardet.horini, pro_controltardet.horfin, pro_controltardet.cant, mae_unidades.abrev, pro_controltardet.observacion, pro_controltardet.tipo AS idtipo, IIf([pro_controltardet].[tipo]=1,[pla_empleados].[id],[pro_controltardet].[idref]) AS idref, pro_controltardet.idrec, pro_controltardet.idtar, pro_controltardet.idunimed,pro_controltardet.observado,pro_controltardet.reproceso " _
'        + vbCr + " FROM pro_tareas RIGHT JOIN (mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN ((pro_receta RIGHT JOIN pro_controltardet ON pro_receta.id = pro_controltardet.idrec) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON alm_inventario.id = pro_receta.iditem) ON mae_unidades.id = pro_controltardet.idunimed) ON pro_tareas.id = pro_controltardet.idtar " _
'        + vbCr + " WHERE (((pro_controltardet.idctr)=" & RstFrm("id") & ")) " _
'        + vbCr + " ORDER BY pro_controltardet.tipo, IIf([pro_controltardet].[tipo]=1,[pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom],'GRUPO Nº ' & [pro_controltardet].[idref]),pro_controltardet.horini, alm_inventario.descripcion, pro_tareas.abrev; "
'
'    RST_Busq RstTmp, nSQL, xCon
'    DoEvents
'    If RstTmp.RecordCount <> 0 Then
'        DoEvents
'        Agregando = True
'        With Fg1
'            .Rows = 1
'            RstTmp.MoveFirst
'            Do While Not RstTmp.EOF
'                DoEvents
'                .Rows = .Rows + 1
'                .TextMatrix(.Rows - 1, 1) = NulosC(RstTmp.Fields("numlote"))
'                .TextMatrix(.Rows - 1, 2) = NulosN(RstTmp.Fields("tipo"))
'                .TextMatrix(.Rows - 1, 3) = NulosC(RstTmp.Fields("nombres"))
'                .TextMatrix(.Rows - 1, 4) = NulosC(RstTmp.Fields("tarea"))
'                .TextMatrix(.Rows - 1, 5) = NulosC(RstTmp.Fields("producto"))
'                If IsDate(RstTmp.Fields("horini")) = True Then .TextMatrix(.Rows - 1, 6) = Format(RstTmp.Fields("horini"), FORMAT_HORA_SIN_SEGUNDO)
'                If IsDate(RstTmp.Fields("horfin")) = True Then .TextMatrix(.Rows - 1, 7) = Format(RstTmp.Fields("horfin"), FORMAT_HORA_SIN_SEGUNDO)
'                .TextMatrix(.Rows - 1, 8) = NulosN(RstTmp.Fields("cant"))
'                .TextMatrix(.Rows - 1, 9) = NulosC(RstTmp.Fields("abrev"))
'                .TextMatrix(.Rows - 1, 10) = NulosC(RstTmp.Fields("observacion"))
'
'                .TextMatrix(.Rows - 1, 11) = NulosN(RstTmp.Fields("idtipo"))
'                .TextMatrix(.Rows - 1, 12) = NulosN(RstTmp.Fields("idref"))
'                .TextMatrix(.Rows - 1, 13) = NulosN(RstTmp.Fields("idrec"))
'                .TextMatrix(.Rows - 1, 14) = NulosN(RstTmp.Fields("idtar"))
'                .TextMatrix(.Rows - 1, 15) = NulosN(RstTmp.Fields("idunimed"))
'
'                .TextMatrix(.Rows - 1, 16) = NulosN(RstTmp.Fields("corr"))
'
'                .TextMatrix(.Rows - 1, 17) = NulosN(RstTmp.Fields("cant"))
'
'                .TextMatrix(.Rows - 1, 18) = NulosN(RstTmp.Fields("observado"))
'
'                .TextMatrix(.Rows - 1, 19) = NulosN(RstTmp.Fields("reproceso"))
'
'                '---
'                RstTmp.MoveNext
'            Loop
'        End With
'
'    End If
'
'    '--cargar datos de los grupos
'    pCargarDatosRstTemp 0, NulosN(RstFrm("id")), 0, True
'    '--cargar datos de las taras vs peso
'    pCargarDatosRstTemp 1, NulosN(RstFrm("id")), 0, True
'    '--------------------------------------------
'    Set RstTmp = Nothing
'    GRID_AGRUPAR Fg1, 3
'    Agregando = False
'    Me.MousePointer = vbDefault
'    Exit Sub
'error:
'    SHOW_ERROR Me.Name, "MuestraDetalle"
'    Me.MousePointer = vbDefault
'    Set RstTmp = Nothing
'    Agregando = False
'End Sub
'
'Private Sub pHabilitarObj(band As Boolean)
'    habilitar_Locked TxtFecha, Not band
'    habilitar_Locked txt_cb, Not band
'    habilitar Me.cb, band
'    habilitar cmd, band
'    Unload FrmControlTareaGr
'End Sub
'
'Sub Blanquea()
'    LimpiaText TxtFecha
'    LimpiaText txt
'    LimpiaText txt_cb
'    LimpiaText lbl_cod
'    LimpiaText lbl_cb
'    LimpiaText lblPesoTara
'    Fg1.Rows = Fg1.FixedRows
'    Set RstGrDet = Nothing
'    Set RstGrDetTara = Nothing
'    FraTarea.Visible = False
'    FraEditor.Visible = False
'End Sub
'
'Sub ActivaTool()
'    Dim a&
'    For a = 1 To Toolbar1.Buttons.Count
'        Toolbar1.Buttons(a).Enabled = Not Toolbar1.Buttons(a).Enabled
'    Next a
'End Sub
'
'Private Sub Nuevo()
'    QueHace = 1
'    mRowAdd = -999
'    '-----------------------------------
'    TabOne1.CurrTab = 1
'
'    TabOne1.TabEnabled(0) = False
'    ActivaTool
'    TxtFecha(0).Valor = Date
'    Blanquea
'    pHabilitarObj True
'    Label1.Caption = "Agregando Seguimiento de Tarea"
'
'    TxtFecha(0).Enabled = True
'    TxtFecha(0).SetFocus
'    pConfigurarGrilla
'    Fg1.SelectionMode = flexSelectionFree
'    '--agregando un registro por defecto
'    Fg1.Rows = 2
'    Fg1.TextMatrix(Fg1.Rows - 1, 16) = mRowAdd '--codigo de inicio
'    '--cargar el temporal a la tara
'    pCargarDatosRstTemp 1, -10, 0
'End Sub
'
'Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
'    If ButtonMenu.Index = 1 Then pImprimir True
'
'    If ButtonMenu.Index = 2 Then pImprimir
'
'End Sub
'
'Function Grabar() As Boolean
'    If fValidarDatos() = False Then Exit Function
'
'    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " el Seguimiento de Tareas", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo SALIR
'
'
'    Dim RstCab As New ADODB.Recordset
'    Dim RstDet As New ADODB.Recordset
'    Dim RstGr As New ADODB.Recordset
'    Dim RstDetTara As New ADODB.Recordset
'    Dim RstGrTara As New ADODB.Recordset
'    Dim xCod&, xCol&, xFil&, xItem&
'
''    On Error GoTo LaCague
'
'    xCon.BeginTrans
'    Me.MousePointer = vbHourglass
'    If QueHace = 1 Then
'        RST_Busq RstCab, "SELECT top 1 * FROM pro_controltar ", xCon
'        xCod = HallaCodigoTabla("pro_controltar", xCon, "id")
'        RstCab.AddNew
'        RstCab("id") = xCod
'    Else
'        xCod = RstFrm("id")
'        RST_Busq RstCab, "SELECT * FROM pro_controltar WHERE id =" & xCod & "", xCon
'        xCon.Execute "DELETE * FROM pro_controltardetgrpes WHERE idctr = " & xCod & ""
'        xCon.Execute "DELETE * FROM pro_controltardetgr WHERE idctr = " & xCod & ""
'        xCon.Execute "DELETE * FROM pro_controltardetpes WHERE idctr = " & xCod & ""
'        xCon.Execute "DELETE * FROM pro_controltardet WHERE idctr = " & xCod & ""
'
'
'    End If
'    '------------------------
'    mIdRegistro = xCod
'    '------------------------
'    RST_Busq RstDet, "SELECT top 1 * FROM pro_controltardet", xCon
'    RST_Busq RstDetTara, "SELECT top 1 * FROM pro_controltardetpes", xCon
'    RST_Busq RstGr, "SELECT top 1 * FROM pro_controltardetgr", xCon
'    RST_Busq RstGrTara, "SELECT top 1 * FROM pro_controltardetgrpes", xCon
'
'
'    RstCab("fchtra") = CDate(TxtFecha(0).Valor)
'    RstCab("idarea") = NulosN(lbl_cod(0).Caption)
'    RstCab("idres") = NulosN(lbl_cod(1).Caption)
'
'    RstCab.Update
'
'    For xFil = 1 To Fg1.Rows - 1
'
'        RstDet.AddNew
'
'        '--codigo
'        RstDet("idctr") = xCod
'        RstDet("corr") = xFil
'        '---
'        RstDet("numlote") = NulosC(Fg1.TextMatrix(xFil, 1))
'        RstDet("tipo") = NulosN(Fg1.TextMatrix(xFil, 11))
'        RstDet("idref") = NulosN(Fg1.TextMatrix(xFil, 12))
'        RstDet("idrec") = NulosN(Fg1.TextMatrix(xFil, 13))
'        RstDet("idtar") = NulosN(Fg1.TextMatrix(xFil, 14))
'        If IsDate(Fg1.TextMatrix(xFil, 6)) = True Then RstDet("horini") = CDate(Fg1.TextMatrix(xFil, 6))
'        If IsDate(Fg1.TextMatrix(xFil, 7)) = True Then RstDet("horfin") = CDate(Fg1.TextMatrix(xFil, 7))
'        RstDet("cant") = NulosN(Fg1.TextMatrix(xFil, 8))
'        RstDet("idunimed") = NulosN(Fg1.TextMatrix(xFil, 15))
'        RstDet("observacion") = NulosC(Fg1.TextMatrix(xFil, 10))
'        RstDet("observado") = NulosN(Fg1.TextMatrix(xFil, 18))
'        RstDet("reproceso") = NulosN(Fg1.TextMatrix(xFil, 19))
'
'        RstDet.Update
'        '***********************************************************************************************************
'        '---------------------------------------------------------------------------------------------------
'        '--registro de taras
'        RstGrDetTara.Filter = "codigo= " & NulosN(Fg1.TextMatrix(xFil, 16)) & " and tipo=0 and idemp =" & NulosN(Fg1.TextMatrix(xFil, 16))
'        If RstGrDetTara.RecordCount <> 0 Then RstGrDetTara.MoveFirst
'        xItem = 1
'        Do While Not RstGrDetTara.EOF
'            RstDetTara.AddNew
'            '--codigo
'            RstDetTara("idctr") = xCod
'            RstDetTara("corr") = xFil
'            RstDetTara("item") = xItem
'            '--fin codigo
'            RstDetTara("idpeso") = NulosN(RstGrDetTara.Fields("idpeso"))
'            RstDetTara("pesouni") = NulosN(RstGrDetTara.Fields("pesouni"))
'            RstDetTara("pesonet") = NulosN(RstGrDetTara.Fields("pesonet"))
'            RstDetTara("cantidad") = NulosN(RstGrDetTara.Fields("cantidad"))
'            RstDetTara("pesotara") = NulosN(RstGrDetTara.Fields("pesotara"))
'            RstDetTara("pesobrut") = NulosN(RstGrDetTara.Fields("pesobrut"))
'            RstDetTara.Update
'            RstGrDetTara.MoveNext
'            xItem = xItem + 1
'        Loop
'        '---------------------------------------------------------------------------------------------------
'        '***********************************************************************************************************
'        '--grabar si es grupo
'        If NulosN(Fg1.TextMatrix(xFil, 11)) = 2 Then
'            RstGrDet.Filter = "codigo= " & NulosN(Fg1.TextMatrix(xFil, 16))
'            If RstGrDet.RecordCount > 0 Then
'                RstGrDet.MoveFirst
'                Do While Not RstGrDet.EOF
'                    RstGr.AddNew
'                    '--codigo
'                    RstGr("idctr") = xCod
'                    RstGr("corr") = xFil
'                    RstGr("idper") = NulosN(RstGrDet.Fields("idemp"))
'                    '-fin codigo
'                    RstGr("cant") = NulosN(RstGrDet.Fields("cant"))
'                    RstGr("cantbrut") = NulosN(RstGrDet.Fields("cantbrut"))
'                    RstGr("activo") = NulosN(RstGrDet.Fields("activo"))
'
'                    If IsDate(RstGrDet.Fields("horini")) = True Then RstGr("horini") = CDate(RstGrDet.Fields("horini"))
'                    If IsDate(RstGrDet.Fields("horfin")) = True Then RstGr("horfin") = CDate(RstGrDet.Fields("horfin"))
'
'                    RstGr.Update
'
'                    '---------------------------------------------------------------------------------------------------
'                    If NulosN(RstGrDet.Fields("activo")) = -1 Then '--solo los activos
'                        '--registro de taras
'                        RstGrDetTara.Filter = "codigo= " & NulosN(Fg1.TextMatrix(xFil, 16)) & " and tipo=1 and idemp =" & NulosN(RstGrDet.Fields("idemp"))
'                        If RstGrDetTara.RecordCount <> 0 Then RstGrDetTara.MoveFirst
'                        xItem = 1
'                        Do While Not RstGrDetTara.EOF
'                            If NulosN(RstGrDetTara.Fields("pesobrut")) <> 0 And NulosN(RstGrDetTara.Fields("pesonet")) <> 0 Then
'                                RstGrTara.AddNew
'                                '--codigo
'                                RstGrTara("idctr") = xCod
'                                RstGrTara("corr") = xFil
'                                RstGrTara("idper") = NulosN(RstGrDet.Fields("idemp"))
'                                RstGrTara("item") = xItem
'                                '--fin codigo
'                                RstGrTara("idpeso") = NulosN(RstGrDetTara.Fields("idpeso"))
'                                RstGrTara("pesouni") = NulosN(RstGrDetTara.Fields("pesouni"))
'                                RstGrTara("pesonet") = NulosN(RstGrDetTara.Fields("pesonet"))
'                                RstGrTara("cantidad") = NulosN(RstGrDetTara.Fields("cantidad"))
'                                RstGrTara("pesotara") = NulosN(RstGrDetTara.Fields("pesotara"))
'                                RstGrTara("pesobrut") = NulosN(RstGrDetTara.Fields("pesobrut"))
'                                RstGrTara.Update
'                                xItem = xItem + 1
'                            End If
'                            RstGrDetTara.MoveNext
'                        Loop
'                    End If
'                    '---------------------------------------------------------------------------------------------------
'
'                    RstGrDet.MoveNext
'                Loop
'            End If
'        End If
'        '***********************************************************************************************************
'    Next xFil
'
'    xCon.CommitTrans
'    MsgBox "El seguimiento de la Tarea se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
'
'    Grabar = True
'SALIR:
'    Me.MousePointer = vbDefault
'    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstGr = Nothing:
'    Exit Function
'LaCague:
'    xCon.RollbackTrans
'    Me.MousePointer = vbDefault
'    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstGr = Nothing:
'    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
'    Grabar = False
'End Function
'
'Private Function fValidarDatos() As Boolean
'    If TxtFecha(0).Valor = "" Or IsDate(TxtFecha(0).Valor) = False Then
'        MsgBox "No ha especificado la fecha de la Programación ", vbInformation, xTitulo
'        TxtFecha(0).SetFocus
'        Exit Function
'    End If
'
'    Dim band As Integer
'    band = Validar(txt_cb)
''    If band <> -1 Then
''       MsgBox "Llene el Campo de " & lbl_cb_capt(band).Caption, vbInformation, xTitulo
''       txt_cb(band).SetFocus
''       Exit Function
''    End If
'
'
'    If Fg1.Rows = 1 Then
'        MsgBox "No ha especificado el registro de las tareas", vbInformation, xTitulo
'        Fg1.SetFocus
'        Exit Function
'    End If
'    '---------------------------------------------------------------------------
'    '--validar la grilla
'    Dim mRow&, mCol&
'
'    mCol = -1
'    For mRow = 1 To Fg1.Rows - 1
'        If NulosN(Fg1.TextMatrix(mRow, 2)) = 0 Then '--tipo de persona
'            MsgBox "Seleccione el tipo Individual / Grupal", vbExclamation, xTitulo
'            mCol = 2
'        ElseIf NulosN(Fg1.TextMatrix(mRow, 12)) = 0 Then '--persona/grupo
'            MsgBox "Seleccione el " & IIf(NulosN(Fg1.TextMatrix(mRow, 2)) = 1, "Personal", "Nº de Grupo"), vbExclamation, xTitulo
'            mCol = 3
''        ElseIf NulosN(Fg1.TextMatrix(mRow, 14)) = 0 Then '--tarea
''            If NulosN(Fg1.TextMatrix(mRow, 13)) = 0 Then '--producto
''                MsgBox "Seleccione el Producto; Si la tarea es diversa, ingrese sólo la tarea", vbExclamation, xTitulo
''                mCol = 4
''            Else
''                MsgBox "Seleccione la Tarea", vbExclamation, xTitulo
''                mCol = 5
''            End If
'''        ElseIf IsDate(Fg1.TextMatrix(mRow, 6)) = False Then '--hora ini
'''            MsgBox "Falta ingresar la Hora de Inicio", vbExclamation, xTitulo
'''            mCol = 6
'''        ElseIf IsDate(Fg1.TextMatrix(mRow, 7)) = False Then '--hora fin
'''            MsgBox "Falta ingresar la Hora Final", vbExclamation, xTitulo
'''            mCol = 7
'''        ElseIf NulosN(Fg1.TextMatrix(mRow, 8)) = 0 Then '--cantidad
'''            MsgBox "Falta ingresar la Cantidad", vbExclamation, xTitulo
'''            mCol = 8
'        ElseIf NulosN(Fg1.TextMatrix(mRow, 15)) = 0 Then '--unidad de medida
'            MsgBox "Falta ingresar la unidad de medida", vbExclamation, xTitulo
'            mCol = 9
'        End If
'
'        If mCol <> -1 Then Exit For
'    Next mRow
'    If mCol <> -1 Then
'        Agregando = True:  Fg1.Row = mRow: Fg1.Col = mCol: Agregando = False
'        Fg1.SetFocus
'        Exit Function
'    End If
'    '---------------------------------------------------------------------------
'
'    fValidarDatos = True
'End Function
'
'Private Sub pCargarGrid()
'    On Error GoTo error
'    Dim nSQL  As String
'
'    lblperiodo(0).Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
'    lblperiodo(1).Caption = lblperiodo(0).Caption
'
'    nSQL = "SELECT pro_controltar.id, pro_controltar.fchtra, pro_controltar.idarea, mae_area.descripcion AS area, pro_controltar.idres, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS encargado,pla_empleados.numdoc " _
'        + vbCr + " FROM mae_area RIGHT JOIN (pro_controltar LEFT JOIN (pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) ON pro_controltar.idres = pro_emp.id) ON mae_area.id = pro_controltar.idarea " _
'        + vbCr + " WHERE (((Year([pro_controltar].[fchtra]))=" & AnoTra & ") AND ((Month([pro_controltar].[fchtra]))=" & mMesActivo & ")) " _
'        + vbCr + " ORDER BY pro_controltar.fchtra,mae_area.descripcion "
'
'
'    Me.MousePointer = vbHourglass
'    RST_Busq RstFrm, nSQL, xCon
'
'    Set Dg3.DataSource = RstFrm
'    Me.MousePointer = vbDefault
'Exit Sub
'error:
'    Me.MousePointer = vbDefault
'    SHOW_ERROR Me.Name, "pCargarGrid"
'End Sub
'
'Private Sub CambiarMes()
'    mMesActivo = SeleccionaMes(xCon)
'    TabOne1.CurrTab = 0
'    If mMesActivo = 0 Or mMesActivo = 13 Then
'        MsgBox "Selecione un Periodo Correcto", vbExclamation, xTitulo
'        CambiarMes
'        Exit Sub
'    End If
'    pCargarGrid
'End Sub
'
'Private Sub pImprimir(Optional IMP_LISTADO As Boolean = False)
'
'    On Error GoTo error
'
'    Me.MousePointer = vbHourglass
'    If IMP_LISTADO = False Then
'        If Me.TabOne1.CurrTab = 0 Then
'
'        Else
'''            MsgBox "Primero muestre el detalle del Registro" + vbCr + _
'''                   "Luego inténtelo otra vez", vbExclamation, xTitulo
'        End If
'    Else
'
'        TDB_IMPRIMIR Dg3, "IMPRESIÓN DE PRODUCCIÓN", "LISTADO DE PRODUCCIÓN  -  Periodo: " + MonthName(mMesActivo, False)
'
'    End If
'
'    Me.MousePointer = vbDefault
'    Exit Sub
'error:
'    Me.MousePointer = vbDefault
'    SHOW_ERROR Me.Name, "pImprimir"
'
'End Sub
'
'Sub Buscar()
'    On Error GoTo error
'    TabOne1.CurrTab = 0
'
'    Dim RstTmp As New ADODB.Recordset
'    Dim nSQL As String
'
'    Dim xCampos(3, 4) As String
'
'    xCampos(0, 0) = "Fch.Trab":         xCampos(0, 1) = "fchtra":     xCampos(0, 2) = "1000":    xCampos(0, 3) = "F"
'    xCampos(1, 0) = "Area":             xCampos(1, 1) = "area":       xCampos(1, 2) = "1500":   xCampos(1, 3) = "C"
'    xCampos(2, 0) = "Respoinsable":     xCampos(2, 1) = "encargado":  xCampos(2, 2) = "3500":   xCampos(2, 3) = "C"
'
'    nSQL = "SELECT pro_controltar.id, pro_controltar.fchtra, pro_controltar.idarea, mae_area.descripcion AS area, pro_controltar.idres, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS encargado " _
'        + vbCr + " FROM mae_area RIGHT JOIN (pro_controltar LEFT JOIN (pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) ON pro_controltar.idres = pro_emp.id) ON mae_area.id = pro_controltar.idarea " _
'        + vbCr + " WHERE (((Year([pro_controltar].[fchtra]))=" & AnoTra & ") AND ((Month([pro_controltar].[fchtra]))=" & mMesActivo & "));"
'
'    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), "Buscando Area", "fchtra", "fchtra", Principio
'    If RstTmp.State = 0 Then GoTo SALIR
'    If RstTmp.EOF = True And RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR
'
'    RstFrm.MoveFirst
'    RstFrm.Find "id = " + CStr(RstTmp("id"))
'SALIR:
'    Set RstTmp = Nothing
'    Exit Sub
'error:
'    Set RstTmp = Nothing
'    SHOW_ERROR Me.Name, "Buscar"
'End Sub
'
'Private Sub Filtrar()
'
'    Dim xCampos(2, 4) As String
'    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
'
'    xCampos(0, 0) = "Fch.Trab":         xCampos(0, 1) = "fchtra":     xCampos(0, 2) = "F":         xCampos(0, 3) = "800"
'    xCampos(1, 0) = "Area":             xCampos(1, 1) = "area":       xCampos(1, 2) = "C":         xCampos(1, 3) = "1000"
'    xCampos(2, 0) = "Respoinsable":     xCampos(2, 1) = "encargado":  xCampos(2, 2) = "C":         xCampos(2, 3) = "1500"
'
'    TabOne1.CurrTab = 0
'
'    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg3
'
'
'End Sub
'
'Private Sub pConfigurarGrilla()
'
'    Dim tFormat$
'    Agregando = True
'    With Fg1 '--de los ingredientes
'        .Rows = 1
'        .Cols = 20
'        .FixedRows = 1
'        .RowHeight(0) = 250
'        .FrozenCols = 5
'        .TextMatrix(0, 1) = "Nº Lote":              .ColWidth(1) = 1200:     .ColAlignment(1) = flexAlignLeftCenter:    .Row = 0: .Col = 1: .CellAlignment = flexAlignLeftCenter
'        .TextMatrix(0, 2) = "Tipo":                 .ColWidth(2) = 600:     .ColAlignment(2) = flexAlignLeftCenter:    .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftCenter
'        .TextMatrix(0, 3) = "Personal/Nº Grupo":    .ColWidth(3) = 1550:    .ColAlignment(3) = flexAlignLeftCenter:    .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftCenter
'        .TextMatrix(0, 4) = "Tarea":                .ColWidth(4) = 2000:    .ColAlignment(4) = flexAlignLeftCenter:    .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftCenter
'        .TextMatrix(0, 5) = "Producto":             .ColWidth(5) = 2400:    .ColAlignment(5) = flexAlignLeftCenter:    .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftCenter
'        .TextMatrix(0, 6) = "H.Inicio":             .ColWidth(6) = 800:     .ColAlignment(6) = flexAlignCenterCenter:  .Row = 0: .Col = 6: .CellAlignment = flexAlignCenterCenter
'        .TextMatrix(0, 7) = "H.Final":              .ColWidth(7) = 800:     .ColAlignment(7) = flexAlignCenterCenter:  .Row = 0: .Col = 7: .CellAlignment = flexAlignCenterCenter
'        .TextMatrix(0, 8) = "Cant":                 .ColWidth(8) = 850:     .ColAlignment(8) = flexAlignRightCenter:   .Row = 0: .Col = 8: .CellAlignment = flexAlignRightCenter
'        .TextMatrix(0, 9) = "U.M.":                 .ColWidth(9) = 500:     .ColAlignment(9) = flexAlignCenterCenter:  .Row = 0: .Col = 9: .CellAlignment = flexAlignCenterCenter
'        .TextMatrix(0, 10) = "Observacion":         .ColWidth(10) = 1000:   .ColAlignment(10) = flexAlignLeftCenter:   .Row = 0: .Col = 10: .CellAlignment = flexAlignLeftCenter
'
'        .TextMatrix(0, 11) = "IdTipo":          .ColWidth(11) = 0:
'        .TextMatrix(0, 12) = "IdRef":           .ColWidth(12) = 0:
'        .TextMatrix(0, 13) = "IdRec":           .ColWidth(13) = 0:
'        .TextMatrix(0, 14) = "IdTar":           .ColWidth(14) = 0:
'        .TextMatrix(0, 15) = "IdUnimed":        .ColWidth(15) = 0:
'
'        .TextMatrix(0, 16) = "Codigo":          .ColWidth(16) = 0:
'        .TextMatrix(0, 17) = "CantTmp":         .ColWidth(17) = 0: '--su uso sera para el calculo automatico
'        .TextMatrix(0, 18) = "Obs":             .ColWidth(18) = 700:   .ColAlignment(18) = flexAlignCenterCenter:   .Row = 0: .Col = 18: .CellAlignment = flexAlignCenterCenter
'        .TextMatrix(0, 19) = "Reproceso":     .ColWidth(19) = 855:   .ColAlignment(18) = flexAlignCenterCenter:   .Row = 0: .Col = 19: .CellAlignment = flexAlignCenterCenter
'
''        .ColFormat(6) = FORMAT_HORA_LARGO
''        .ColFormat(7) = FORMAT_HORA_LARGO
'        .ColEditMask(6) = "##:##" '--hora inicio
'        .ColEditMask(7) = "##:##" '--hora fin
'        .ColFormat(8) = FORMAT_MONTO '--cantidad
'
'        .SelectionMode = flexSelectionByRow
'
'        GRID_COMBOLIST Fg1, 3 '--persona / grupo
'        GRID_COMBOLIST Fg1, 4 '--tarea
'        GRID_COMBOLIST Fg1, 5 '--producto
'
'        GRID_COMBOLIST Fg1, 8 '--cantidad
'
'        GRID_COMBOLIST Fg1, 9 '--unidad de medida
'
'        '--Tipo de Origen (Materia Prima; Producto)
'        .ColComboList(2) = "#1;Individual|#2;Grupal"
'
'        .ColDataType(18) = flexDTBoolean
'        .ColDataType(19) = flexDTBoolean
'
'        DoEvents
'
'    End With
'
'    fg(1).ColWidth(6) = 0
'    fg(1).ColWidth(7) = 0
'    GRID_COMBOLIST fg(1), 3 '--peso - tara
'
'    fg(0).ColWidth(3) = 0 '--idrec
'
'    Agregando = False
'End Sub
'
'
''*******************************************************************************************
'
'Private Sub cb_Click(Index As Integer)
'    If QueHace = 3 Then Exit Sub
'    Dim xCampos() As String
'    Dim nCampoBusca As String
'    Dim nSQL As String
'    Dim nTitulo As String
'    On Error GoTo error
'    Select Case Index
'        Case 0 '--area
'            ReDim xCampos(2, 3) As String
'            xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
'            xCampos(1, 0) = "Id":           xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
'
'            nTitulo = "Buscando Area"
'            nSQL = "SELECT mae_area.id, mae_area.descripcion AS nombre, mae_area.id AS cod, mae_area.id AS idarea, pro_emp.id AS idper, pla_empleados.id AS idemp, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS encargado, pla_empleados.numdoc " _
'                + vbCr + " FROM ((pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) RIGHT JOIN pro_area ON pro_emp.id = pro_area.idper) INNER JOIN mae_area ON pro_area.idarea = mae_area.id "
'
'        Case 1 '--responsable de area
'            ReDim xCampos(2, 3) As String
'            xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
'            xCampos(1, 0) = "Id":           xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
'
'            nTitulo = "Buscando Responsable de Area "
'            nSQL = "SELECT pla_empleados.numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, pro_emp.id AS cod, pla_empleados.id AS idemp, mae_dociden.abrev " _
'                + vbCr + " FROM mae_dociden RIGHT JOIN ((pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) ON mae_dociden.id = pla_empleados.idtipdoc " _
'                + vbCr + " WHERE (((pro_empdet.idfun)=5)); "
'
'        Case 2 '--perdida de peso
'            ReDim xCampos(3, 3) As String
'            xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "2500":   xCampos(0, 3) = "C"
'            xCampos(1, 0) = "Peso":         xCampos(1, 1) = "peso":      xCampos(1, 2) = "900":    xCampos(1, 3) = "N"
'            xCampos(2, 0) = "Und Destino":  xCampos(2, 1) = "destabrev": xCampos(2, 2) = "1500":   xCampos(2, 3) = "C"
'            nTitulo = "Buscando Contenedor"
'
'            nSQL = "SELECT pro_pesotara.id,  '1 ' & [mae_unidades].[abrev] & ' => ' & [pro_pesotara].[peso] & ' ' & [mae_unidades_1].[abrev] AS ref, pro_pesotara.id AS cod, pro_pesotara.descripcion as nombre,pro_pesotara.abrev, pro_pesotara.peso, mae_unidades_1.abrev AS destabrev  " _
'                + vbCr + " FROM mae_unidades AS mae_unidades_1 INNER JOIN (mae_unidades INNER JOIN pro_pesotara ON mae_unidades.id = pro_pesotara.idundori) ON mae_unidades_1.id = pro_pesotara.idunddes;"
'
'        Case 3 '--buscar las tareas relacionados a productos
'            If optTarea(0).Value = True Then
'                ReDim xCampos(2, 3) As String
'                xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
'                xCampos(1, 0) = "Id":           xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
'
'                nTitulo = "Buscando Tareas"
'                nSQL = "SELECT pro_tareas.id, pro_tareas.descripcion as nombre, pro_tareas.id as cod FROM pro_tareas WHERE (((pro_tareas.diverso)=0)); "
'            Else
'                ReDim xCampos(3, 4) As String
'                xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "codpro":   xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
'                xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "nombre":   xCampos(1, 2) = "5000":    xCampos(1, 3) = "C"
'                xCampos(2, 0) = "CodReceta":    xCampos(2, 1) = "codrec":   xCampos(2, 2) = "1200":    xCampos(2, 3) = "C"
'
'                nTitulo = "Buscando Recetas"
'                nSQL = "SELECT pro_receta.codrec, alm_inventario.descripcion AS nombre, pro_receta.id AS cod, alm_inventario.codpro " _
'                    + vbCr + " FROM alm_inventario INNER JOIN (pro_receta INNER JOIN pro_recetatar ON pro_receta.id = pro_recetatar.idrec) ON alm_inventario.id = pro_receta.iditem " _
'                    + vbCr + " GROUP BY pro_receta.codrec, alm_inventario.descripcion, pro_receta.id,alm_inventario.codpro "
'            End If
'    End Select
'    If Index <> 2 Then
'
'    Else
'
'    End If
'    Dim RstTmp As New ADODB.Recordset
'    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
'
'    If RstTmp.State = 0 Then GoTo SALIR
'    If RstTmp.EOF = True Or RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR
'
'    lbl_cod(Index).Tag = lbl_cod(Index).Caption
'
'    txt_cb(Index).Text = NulosC(RstTmp.Fields(0))  '--TEXTO A MOSTRAR
'    lbl_cb(Index).Caption = NulosC(RstTmp.Fields(1)) '--NOMBRE
'    lbl_cod(Index).Caption = NulosN(RstTmp.Fields(2)) '--CODIGO
'    lbl_cb(Index).ToolTipText = NulosC(RstTmp.Fields(1))  '--NOMBRE
'
'    Select Case Index
'        Case 0 '--area
'            '--poner datos del encargado por defecto
'            txt_cb(1).Text = NulosC(RstTmp.Fields("numdoc"))  '--TEXTO A MOSTRAR
'            lbl_cb(1).Caption = NulosC(RstTmp.Fields("encargado")) '--NOMBRE
'            lbl_cod(1).Caption = NulosN(RstTmp.Fields("idper")) '--CODIGO
'            lbl_cb(1).ToolTipText = NulosC(RstTmp.Fields("encargado"))  '--NOMBRE
'
'            If NulosN(lbl_cod(1).Caption) = 0 Then txt_cb(1).SetFocus
'
'        Case 1 '--encargado
'            If Fg1.Rows > Fg1.FixedRows Then
'                Fg1.Col = 1:    Fg1.Row = Fg1.Rows - 1:   Fg1.SetFocus
'            Else
'                cmd(0).SetFocus
'            End If
'
'        Case 2 '--perdida de peso
'            lblPesoTara(0).Caption = NulosN(RstTmp("peso"))
'            lblPesoTara(1).Caption = NulosC(RstTmp("abrev"))
'
'            txt_cb(0).SetFocus
'        Case 3 '--tareas relacionadas a productos
'            If NulosN(lbl_cod(3).Caption) <> 0 Then pCargarTareasReceta
'    End Select
'SALIR:
'    Set RstTmp = Nothing
'Exit Sub
'error:
'    Me.MousePointer = vbDefault
'    Set RstTmp = Nothing
'    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
'End Sub
'
'
'Private Sub txt_cb_Change(Index As Integer)
'    If QueHace = 3 Then Exit Sub
'    If txt_cb(Index).Text = "" Then
'        Me.lbl_cb(Index).Caption = ""
'        Me.lbl_cod(Index).Caption = ""
'    End If
'    If Index = 0 Then txt_cb(1).Text = ""
'    If Index = 2 Then LimpiaText lblPesoTara
'    If Index = 3 Then fg(0).Rows = 1
'
'End Sub
'
'Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If QueHace = 3 Then Exit Sub
'    If txt_cb(Index).Locked = True Then Exit Sub
'    If KeyCode = vbKeyF5 Then
'        cb_Click Index
'        Exit Sub
'    End If
'End Sub
'
'Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If Index <> 1 Then
'            SendKeys vbTab
'        Else
'            If Fg1.Rows >= 2 Then
'                Fg1.Row = 1: Fg1.Col = 1
'            Else
'                Fg1.Row = Fg1.Rows - 1: Fg1.Col = 1
'            End If
'            Fg1.SetFocus
'        End If
'        Exit Sub
'    End If
'    If validar_numero(KeyAscii) = False Then KeyAscii = 0
'End Sub
'
'Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)
'    If QueHace = 3 Then Exit Sub
'    If txt_cb(Index).Text = "" Then Exit Sub
'
'    Dim RstTmp As New ADODB.Recordset
'    Dim nSQL As String
'    On Error GoTo error
'    Select Case Index
'        Case 0 '--area
'            nSQL = "SELECT mae_area.id, mae_area.descripcion AS nombre, mae_area.id AS cod, mae_area.id AS idarea, pro_emp.id AS idper, pla_empleados.id AS idemp, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS encargado, pla_empleados.numdoc " _
'                + vbCr + " FROM ((pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) RIGHT JOIN pro_area ON pro_emp.id = pro_area.idper) INNER JOIN mae_area ON pro_area.idarea = mae_area.id " _
'                + vbCr + " WHERE mae_area.id = " & NulosN(txt_cb(Index).Text) & " ;"
'
'        Case 1 '--encargado de area
'            nSQL = "SELECT pla_empleados.numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, pro_emp.id AS cod, pla_empleados.id AS idemp, mae_dociden.abrev " _
'                + vbCr + " FROM mae_dociden RIGHT JOIN ((pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) ON mae_dociden.id = pla_empleados.idtipdoc " _
'                + vbCr + " WHERE (((pro_empdet.idfun)=5)) and pla_empleados.numdoc = '" & NulosC(txt_cb(Index).Text) & "';"
'
'        Case 2 '--perdida de peso
'            nSQL = "SELECT pro_pesotara.id,  '1 ' & [mae_unidades].[abrev] & ' => ' & [pro_pesotara].[peso] & ' ' & [mae_unidades_1].[abrev] AS ref, pro_pesotara.id AS cod, pro_pesotara.descripcion as nombre,pro_pesotara.abrev, pro_pesotara.peso, mae_unidades_1.abrev AS destabrev  " _
'                + vbCr + " FROM mae_unidades AS mae_unidades_1 INNER JOIN (mae_unidades INNER JOIN pro_pesotara ON mae_unidades.id = pro_pesotara.idundori) ON mae_unidades_1.id = pro_pesotara.idunddes " _
'                + vbCr + " WHERE pro_pesotara.id= " & NulosN(txt_cb(Index).Text) & ";"
'        Case 3 '--tareas relacionadas con receta
'            nSQL = "SELECT pro_tareas.id, pro_tareas.descripcion as nombre,pro_tareas.id as cod " _
'            + vbCr + " FROM pro_tareas " _
'            + vbCr + " WHERE pro_tareas.diverso=0 and pro_tareas.id= " & NulosN(txt_cb(Index).Text) & ";"
'        Case Else
'            Exit Sub
'    End Select
'
'    If xCon.State = 0 Then GoTo SALIR
'    RST_Busq RstTmp, nSQL, xCon
'
'    If RstTmp.State = 0 Then GoTo SALIR
'
'    lbl_cod(Index).Tag = lbl_cod(Index).Caption
'
'    If RstTmp.RecordCount > 0 Then
'        txt_cb(Index).Text = NulosC(RstTmp.Fields(0))   '--TEXTO A MOSTRAR
'        lbl_cb(Index).Caption = NulosC(RstTmp.Fields(1)) '--NOMBRE
'        lbl_cod(Index).Caption = NulosN(RstTmp.Fields(2)) '--CODIGO
'        lbl_cb(Index).ToolTipText = NulosC(RstTmp.Fields(1)) '--NOMBRE
'        If Index = 0 Then
'            '--poner datos del encargado por defecto
'            txt_cb(1).Text = NulosC(RstTmp.Fields("numdoc"))  '--TEXTO A MOSTRAR
'            lbl_cb(1).Caption = NulosC(RstTmp.Fields("encargado")) '--NOMBRE
'            lbl_cod(1).Caption = NulosN(RstTmp.Fields("idper")) '--CODIGO
'            lbl_cb(1).ToolTipText = NulosC(RstTmp.Fields("encargado"))  '--NOMBRE
'        ElseIf Index = 2 Then
'            lblPesoTara(0).Caption = NulosN(RstTmp("peso"))
'            lblPesoTara(1).Caption = NulosC(RstTmp("abrev"))
'        End If
'    Else
'        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cod(Index).Caption = ""
'    End If
'
'    If Index = 3 Then pCargarTareasReceta
'    '--------------
'    Set RstTmp = Nothing
'    Exit Sub
'error:
'    Set RstTmp = Nothing
'    SHOW_ERROR Me.Name, "txt_cb_Validate (" + CStr(Index) + ")"
'    Exit Sub
'SALIR:
'    Set RstTmp = Nothing
'    txt_cb(Index).Text = ""
'End Sub
'
''****************************************************************************************
'
'Private Sub pRegistroAdd()
'    Dim mCol%
'    Dim fInsertar As Boolean
'    If QueHace = 3 Then Exit Sub
'    Agregando = True
'    If Fg1.Rows > Fg1.FixedRows Then
'        If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 2)) = 0 Then '--tipo de persona
'            MsgBox "Seleccione el tipo Individual / Grupal", vbExclamation, xTitulo
'            mCol = 2
'        ElseIf NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 12)) = 0 Then '--persona/grupo
'            MsgBox "Seleccione el " & IIf(NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 2)) = 1, "Personal", "Nº de Grupo"), vbExclamation, xTitulo
'            mCol = 3
''        ElseIf NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 14)) = 0 Then '--tarea
''
''            If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 13)) = 0 Then '--producto
''                MsgBox "Seleccione el Producto; Si la tarea es diversa, ingrese sólo la tarea", vbExclamation, xTitulo
''                mCol = 4
''            Else
''                MsgBox "Seleccione la Tarea", vbExclamation, xTitulo
''                mCol = 5
''            End If
''        ElseIf IsDate(Fg1.TextMatrix(Fg1.Rows - 1, 6)) = False Then '--hora ini
''            MsgBox "Falta ingresar la Hora de Inicio", vbExclamation, xTitulo
''            mCol = 6
''        ElseIf IsDate(Fg1.TextMatrix(Fg1.Rows - 1, 7)) = False Then '--hora fin
''            MsgBox "Falta ingresar la Hora Final", vbExclamation, xTitulo
''            mCol = 7
''        ElseIf NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 8)) = 0 Then '--cantidad
''            MsgBox "Falta ingresar la Cantidad", vbExclamation, xTitulo
''            mCol = 8
'        ElseIf NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 15)) = 0 Then '--unidad de medida
'            MsgBox "Seleccione la unidad de medida", vbExclamation, xTitulo
'            mCol = 9
'
'        Else
'            fInsertar = True
'            mCol = 5
'        End If
'    Else
'        fInsertar = True
'        mCol = 1
'    End If
'
'    If fInsertar = True Then Fg1.AddItem ""
'    If Fg1.Rows > 2 And fInsertar = True Then
'
'        If chkOpcion(1).Value = 1 Then Fg1.TextMatrix(Fg1.Rows - 1, 1) = Fg1.TextMatrix(Fg1.Rows - 2, 1) '--num lote
'
'        If chkOpcion(2).Value = 1 Then
'            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosN(Fg1.TextMatrix(Fg1.Rows - 2, 2)) '--tipo
'            Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosN(Fg1.TextMatrix(Fg1.Rows - 2, 11)) '--idtipo
'            mCol = 3
'        End If
'        If chkOpcion(3).Value = 1 Then
'            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(Fg1.TextMatrix(Fg1.Rows - 2, 3)) '--indivudual/grupal
'            Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosN(Fg1.TextMatrix(Fg1.Rows - 2, 12)) '--cod ref dsfsf
'            mCol = 4
'        End If
'        If chkOpcion(4).Value = 1 Then
'            Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(Fg1.TextMatrix(Fg1.Rows - 2, 5)) '--tarea
'            Fg1.TextMatrix(Fg1.Rows - 1, 14) = NulosN(Fg1.TextMatrix(Fg1.Rows - 2, 14)) '--idtar
'            mCol = 5
'        End If
'        If chkOpcion(5).Value = 1 Then
'            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(Fg1.TextMatrix(Fg1.Rows - 2, 4)) '--producto
'            Fg1.TextMatrix(Fg1.Rows - 1, 13) = NulosN(Fg1.TextMatrix(Fg1.Rows - 2, 13)) '--idrec
'            mCol = 6
'        End If
'
'        If chkOpcion(6).Value = 1 Then
'            Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(Fg1.TextMatrix(Fg1.Rows - 2, 6)) '--hora inicio
'            Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosC(Fg1.TextMatrix(Fg1.Rows - 2, 7)) '--hora fin
'            mCol = 8
'        End If
'
'        Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosC(Fg1.TextMatrix(Fg1.Rows - 2, 9)) '--unidad
'        Fg1.TextMatrix(Fg1.Rows - 1, 15) = NulosN(Fg1.TextMatrix(Fg1.Rows - 2, 15)) '--idunid
'
'        mRowAdd = mRowAdd + 1 '--incrementar
'        Fg1.TextMatrix(Fg1.Rows - 1, 16) = mRowAdd '--codigo
'
'        If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 11)) = 2 Then
'            pCargarDatosRstTemp 0, NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 12)), NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 16)), False
'        End If
'
'    End If
'
'    Fg1.Row = Fg1.Rows - 1
'    Fg1.Col = mCol
'
'    Agregando = False
'
'    Fg1.SetFocus
'
'End Sub
'
'Private Sub pRegistroDel()
'
'    If QueHace = 3 Then Exit Sub
'
'    If Fg1.Row < 1 Then
'        MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Exit Sub
'    End If
'
'    If Fg1.Rows = 1 Then
'        MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Fg1.SetFocus
'        Exit Sub
'    End If
'    If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
'
'    '--descargar el formulario de detalle del grupo
'    Unload FrmControlTareaGr
'    '--eliminar los registros temporales
'    RstRegistroEliminar RstGrDet, "codigo", NulosN(Fg1.TextMatrix(Fg1.Row, 16)), True
'
'    Fg1.RemoveItem Fg1.Row
'
'    If Fg1.Rows > 1 Then
'        Fg1.Row = Fg1.Rows - 1
'        Fg1.Col = 1
'        Fg1.SetFocus
'    Else
'        cmd(0).SetFocus
'    End If
'
'End Sub
'
'Private Sub pCargarDatosRstTemp(mTipo As Integer, idCod1, mRowPosicion, Optional fDesdeBD As Boolean = True)
'    '===================================================================================================
'    'Propósito: Definir la estructura del recordset de los grupos
'    '
'    'Entradas:  mTipo = Especifica el tipo de recordset a definir 0::Grupo; 1::Tara
'    '           idCod1 = codigo del control de tareas
'    '           mRowPosicion = posicion de la fila cuando se seleccione un grupo
'    '           fDesdeBD = false::cuando se esta editando el grid
'    '                      true::consulta directamente de la bd
'    'Resultados: Recordset Definido
'    '===================================================================================================
'
'    Dim RstTmp As New ADODB.Recordset
'    Dim nSQL As String
'    Set RstTmp = Nothing
'
'    '--definir la estructura de recordset
'    If RstGrDet.State = 0 Then
'        nSQL = "SELECT pro_controltardetgr.corr AS codigo, pla_empleados.id AS idemp, pro_controltardet.idref AS idgrupo, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, pro_controltardetgr.cant, pro_controltardetgr.cantbrut, pro_controltardetgr.activo, pro_controltardetgr.horini, pro_controltardetgr.horfin  " _
'            + vbCr + " FROM pla_empleados INNER JOIN (pro_controltardet INNER JOIN pro_controltardetgr ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) ON pla_empleados.id = pro_controltardetgr.idper " _
'            + vbCr + " Where ((pro_controltardetgr.idctr) = -10)"
'        RST_Busq RstTmp, nSQL, xCon
'        DEFINIR_RST_TMP RstGrDet, RstTmp
'    End If
'
'    If mTipo = 0 Then
'
'        '--cargar los datos
'        If fDesdeBD = True Then
'            nSQL = "SELECT pro_controltardetgr.corr AS codigo, pla_empleados.id AS idemp, pro_controltardet.idref AS idgrupo, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, pro_controltardetgr.cant, pro_controltardetgr.cantbrut, pro_controltardetgr.activo, pro_controltardetgr.horini, pro_controltardetgr.horfin " _
'                + vbCr + " FROM pla_empleados INNER JOIN (pro_controltardet INNER JOIN pro_controltardetgr ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) ON pla_empleados.id = pro_controltardetgr.idper " _
'                + vbCr + " WHERE ((pro_controltardetgr.idctr)=" & idCod1 & ") " _
'                + vbCr + " ORDER BY pro_controltardetgr.corr, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom]; "
'
'        Else
'
'            nSQL = "SELECT " & mRowPosicion & " as codigo, pla_empleados.id AS idemp, pro_grupo.id AS idgrupo, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres,0 as cant,0 as cantbrut, -1 AS activo " _
'                + vbCr + " FROM pro_grupo INNER JOIN ((pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_grupodet ON pro_emp.id = pro_grupodet.idper) ON pro_grupo.id = pro_grupodet.idgrupo " _
'                + vbCr + " WHERE (((pro_grupo.id)=" & idCod1 & ")); "
'
'        End If
'    Else
'
'        nSQL = "select * from ( " _
'            + vbCr + " SELECT pro_controltardetgr.corr AS codigo, 1 AS tipo, pro_controltardetgr.idper AS idemp, pro_controltardetgrpes.item, pro_pesotara.abrev, pro_controltardetgrpes.pesouni, pro_controltardetgrpes.pesonet, pro_controltardetgrpes.cantidad, pro_controltardetgrpes.pesotara, pro_controltardetgrpes.pesobrut, pro_controltardetgrpes.idpeso " _
'            + vbCr + " FROM pro_pesotara RIGHT JOIN (pro_controltardetgr INNER JOIN pro_controltardetgrpes ON (pro_controltardetgr.idper = pro_controltardetgrpes.idper) AND (pro_controltardetgr.corr = pro_controltardetgrpes.corr) AND (pro_controltardetgr.idctr = pro_controltardetgrpes.idctr)) ON pro_pesotara.id = pro_controltardetgrpes.idpeso " _
'            + vbCr + " Where (((pro_controltardetgr.idctr) = " & idCod1 & ")) " _
'            + vbCr + " Union " _
'            + vbCr + " SELECT pro_controltardet.corr AS codigo, 0 AS tipo, pro_controltardet.corr AS idemp, pro_controltardetpes.item, pro_pesotara.abrev, pro_controltardetpes.pesouni, pro_controltardetpes.pesonet, pro_controltardetpes.cantidad, pro_controltardetpes.pesotara, pro_controltardetpes.pesobrut, pro_controltardetpes.idpeso " _
'            + vbCr + " FROM pro_controltardet INNER JOIN (pro_pesotara INNER JOIN pro_controltardetpes ON pro_pesotara.id = pro_controltardetpes.idpeso) ON (pro_controltardet.corr = pro_controltardetpes.corr) AND (pro_controltardet.idctr = pro_controltardetpes.idctr) " _
'            + vbCr + " Where (((pro_controltardet.idctr) = " & idCod1 & ")) " _
'            + vbCr + " ) as vw " _
'            + vbCr + " order by vw.codigo, vw.tipo, vw.idemp, vw.item"
'        RST_Busq RstTmp, nSQL, xCon
'
'        If RstGrDetTara.State = 0 Then DEFINIR_RST_TMP RstGrDetTara, RstTmp
'
'    End If
'
'    Set RstTmp = Nothing
'
'    RST_Busq RstTmp, nSQL, xCon
'
'    DoEvents
'    If mTipo = 0 Then
'        If RstTmp.RecordCount <> 0 Then CARGAR_RST_TMP RstGrDet, RstTmp
'    Else
'        If RstTmp.RecordCount <> 0 Then CARGAR_RST_TMP RstGrDetTara, RstTmp
'    End If
'
'    Set RstTmp = Nothing
'
'
'End Sub
'
'
'Private Sub pBuscarVSFlexGrid()
'    On Error GoTo error
'    If Me.TabOne1.CurrTab = 0 Then Exit Sub
'    Dim xExport As New SGI2_funciones.formularios
'    Dim xCampos(3, 3) As String
'    'campo     'columna del grid    'tipo(N:Numerico, C:caracter, F:fecha)      campo predeterminado(0:no se muestra, -1:se muestra al iniciar el formulario)
'    xCampos(0, 0) = "Nº Orden":            xCampos(0, 1) = "1":    xCampos(0, 2) = "C":    xCampos(0, 3) = "0"
'    xCampos(1, 0) = "Personal / Nº Grupo": xCampos(1, 1) = "3":    xCampos(1, 2) = "C":    xCampos(1, 3) = "-1"
'    xCampos(2, 0) = "Producto":            xCampos(2, 1) = "5":    xCampos(2, 2) = "C":    xCampos(2, 3) = "0"
'    xCampos(3, 0) = "Tarea":               xCampos(3, 1) = "4":    xCampos(3, 2) = "C":    xCampos(3, 3) = "0"
'
'    xExport.VSFlexGrid_Buscar Me.hWnd, Fg1, xCampos(), Fg1.Row
'    Set xExport = Nothing
'    Me.MousePointer = vbDefault
'
'    Exit Sub
'error:
'    Me.MousePointer = vbDefault
'    SHOW_ERROR Me.Name, "pBuscarVSFlexGrid"
'End Sub
'
'
'Private Sub pExportarVSFlexGrid()
'    If IsDate(TxtFecha(0).Valor) = False Then
'        MsgBox "Falta especificar la Fecha de Trabajo", vbExclamation, xTitulo
'        TxtFecha(0).SetFocus
'        Exit Sub
'    End If
'
'    If lbl_cod(0).Caption = 0 Then
'        MsgBox "Falta especificar el Area", vbExclamation, xTitulo
'        txt_cb(0).SetFocus
'        Exit Sub
'    End If
'    On Error GoTo error
'
'    Dim oExport As New SGI2_funciones.formularios
'    Dim nTitulo As String
'    Dim nPeriodo As String
'    Dim nTitulo1 As String
'    nTitulo = "Control de Tareas - Area: " & StrConv(lbl_cb(0).Caption, 3)
'    nPeriodo = "Fch. Trabajo: " + TxtFecha(0).Valor
'    nTitulo1 = "Responsable: " & StrConv(lbl_cb(1).Caption, 3)
'
'    Me.MousePointer = vbHourglass
'    oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg1, nTitulo, nPeriodo, nTitulo1, "Control de Tareas"
'    Set oExport = Nothing
'    Me.MousePointer = vbDefault
'    Exit Sub
'error:
'    Me.MousePointer = vbDefault
'    SHOW_ERROR Me.Name, "pExportarVSFlexGrid"
'End Sub
'
'
''******************************************************************************
''******************************************************************************
'
'Private Sub pHabilitarBotonEditor(Origen As Integer, band As Boolean)
'    '===================================================================================================
'    'Propósito: Registrar las cantidades de ya se a una persona o a un grupo
'    '           Ej. peso bruto y cantidades
'    'Entradas:  origen =1::mostrar editor de taras-peso; 2::consultar tarea en recetas; 3::mostrar cuadro de opciones
'    '           band= puede ser true o false
'    '
'    'Resultados: Mostrar/Ocultar las opciones del Ingreso de Datos
'    '
'    '===================================================================================================
'    '--si es true cargar los datos
'    Agregando = True
'    '*********************************************
'    habilitar cmd, Not band
'    habilitar CmdUtil, Not band
'    Fg1.Enabled = Not band
'    Toolbar1.Enabled = Not band
'
'    habilitar cb, Not band
'    habilitar txt_cb, Not band
'
'
'    '*********************************************
'    If Origen = 1 Then '--fra_taras
'
'        FraEditor.Visible = band
'
'       '--true muestra el ingreso de datos
'        If band = True Then
'            If FrmControlTarea.Fg1.Row <= 9 Then
'                FraEditor.Top = 3535
'                FraEditor.Left = 50
'            Else
'                FraEditor.Top = 265
'                FraEditor.Left = 50
'            End If
'            LblTituloFrame.Caption = "Registros: " & Fg1.TextMatrix(Fg1.Row, 3)
'        End If
'
'        If band = True Then
'            fg(1).Rows = 1
'            With RstGrDetTara
'                '--filtrar los registros solo del personal seleccionado
'                .Filter = "codigo = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & " and idemp=" & NulosN(Fg1.TextMatrix(Fg1.Row, 16))
'                If .RecordCount <> 0 Then .MoveFirst
'                Do While Not .EOF
'                    fg(1).Rows = fg(1).Rows + 1
'                    fg(1).TextMatrix(fg(1).Rows - 1, 1) = NulosN(.Fields("pesobrut"))
'                    fg(1).TextMatrix(fg(1).Rows - 1, 2) = NulosN(.Fields("cantidad"))
'                    fg(1).TextMatrix(fg(1).Rows - 1, 3) = NulosC(.Fields("abrev"))
'                    fg(1).TextMatrix(fg(1).Rows - 1, 4) = NulosN(.Fields("pesouni"))
'                    fg(1).TextMatrix(fg(1).Rows - 1, 5) = NulosN(.Fields("pesonet"))
'                    fg(1).TextMatrix(fg(1).Rows - 1, 6) = NulosN(.Fields("idpeso"))
'                    fg(1).TextMatrix(fg(1).Rows - 1, 7) = NulosN(.Fields("item")) '--identificador de fila
'
'                    .MoveNext
'                Loop
'            End With
'            If fg(1).Rows > 1 Then
'                fg(1).Row = fg(1).Rows - 1
'                fg(1).Col = 1
'                fg(1).SetFocus
'            Else
'                CmdEditor(0).SetFocus
'            End If
'
'            lblTotal(1).Caption = Format(GRID_SUMAR_COL(fg(1), 1), FORMAT_MONTO)
'            lblTotal(3).Caption = Format(GRID_SUMAR_COL(fg(1), 5), FORMAT_MONTO)
'
'        Else
'            '--acumular las cantidades
'            If NulosN(lblTotal(0).Caption) <> 0 Then Fg1.TextMatrix(Fg1.Row, 8) = GRID_SUMAR_COL(fg(1), 5)
'            '---------------
'            Fg1.Row = Fg1.Row
'            Fg1.Col = 8
'            Fg1.SetFocus
'        End If
'    ElseIf Origen = 2 Then '--fra_recetas(mostrar)
'
'        FraTarea.Visible = band
'        FraTarea.Enabled = band
'
'        If band = True Then
'
'            If FrmControlTarea.Fg1.Row <= 9 Then
'                FraTarea.Top = 2730
'                FraTarea.Left = 50
'            Else
'                FraTarea.Top = 265
'                FraTarea.Left = 50
'            End If
'
'            fg(0).Rows = 1
'            fg(0).SelectionMode = flexSelectionByRow
'            cb(3).Enabled = True
'            txt_cb(3).Enabled = True
'            txt_cb(3).Locked = False
'            txt_cb(3).Text = ""
'            txt_cb(3).SetFocus
'        End If
'
'    ElseIf Origen = 3 Then '--fra_recetas(mostrar)
'
'        FraOpcion.Visible = band
'        FraOpcion.Enabled = band
'
'        If band = True Then
'            FraOpcion.Top = 3990
'            FraOpcion.Left = 4395
'            chkOpcion(0).SetFocus
'        End If
'
'    End If
'
'    Agregando = False
'End Sub
'
'Private Sub CmdEditor_Click(Index As Integer)
'    Select Case Index
'        Case 0 '--agregar
'            pRegistroAddTara
'        Case 1 '--eliminar
'            pRegistroDelTara
'        Case 2 'cancelar
'            pHabilitarBotonEditor 1, False
'        Case 3 '--aceptar cuadro de opciones
'            pHabilitarBotonEditor 3, False
'    End Select
'End Sub
'
'Private Sub pic_Click(Index As Integer)
'    If Index = 0 Then
'        CmdEditor_Click 2
'    ElseIf Index = 2 Then
'        CmdTarea_Click 0
'    ElseIf Index = 1 Then
'        CmdEditor_Click 3
'    End If
'End Sub
'
'
'Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
'    If Agregando = True Then Exit Sub
'    If QueHace = 3 Then Exit Sub
'    If Index <> 1 Then Exit Sub
'    '------------------------------------------------------------------
'    '--aplicando filtro
'    RstGrDetTara.Filter = "codigo = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & " and tipo=0 and idemp = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16))
'    If RstGrDetTara.RecordCount <> 0 Then RstGrDetTara.MoveFirst
'    RstGrDetTara.Find "item = " & NulosN(fg(1).TextMatrix(Row, 7))
'
'    If RstGrDetTara.EOF = False And RstGrDetTara.BOF = False Then
'
'        RstGrDetTara("cantidad") = NulosN(fg(1).TextMatrix(Row, 2))
'        RstGrDetTara("pesouni") = NulosN(fg(1).TextMatrix(Row, 4))
'
'        RstGrDetTara("pesotara") = NulosN(RstGrDetTara("pesouni")) * NulosN(RstGrDetTara("cantidad"))
'
'        RstGrDetTara("pesobrut") = NulosN(fg(1).TextMatrix(Row, 1))
'
'        RstGrDetTara("pesonet") = NulosN(RstGrDetTara("pesobrut")) - NulosN(RstGrDetTara("pesotara"))
'        '--------------
'        fg(1).TextMatrix(Row, 5) = NulosN(RstGrDetTara("pesonet"))
'
'    End If
'
'    If IsNumeric(fg(1).TextMatrix(Row, Col)) = False Then fg(1).TextMatrix(Row, Col) = 0
'
'    lblTotal(1).Caption = Format(GRID_SUMAR_COL(fg(1), 1), FORMAT_MONTO)
'    lblTotal(3).Caption = Format(GRID_SUMAR_COL(fg(1), 5), FORMAT_MONTO)
'
'    Fg1.TextMatrix(Fg1.Row, 8) = lblTotal(3).Caption
'
'    fActivarAutomaticoCantidad = True
'    Fg1_RowColChange
'    fActivarAutomaticoCantidad = False
'
'    '------------------------------------------------------------------
'
'End Sub
'
'Private Sub Fg_EnterCell(Index As Integer)
'    If QueHace = 3 Or Index <> 1 Then
'        fg(1).Editable = flexEDNone
'        Exit Sub
'    End If
'    fg(1).Editable = flexEDKbdMouse
'End Sub
'
'Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
'    If Index <> 1 Then Exit Sub
'    If Col <> 2 And Col <> 3 And Col <> 5 Then Exit Sub
'    If QueHace = 3 Then Exit Sub
'    Dim xRs As New ADODB.Recordset
'    Dim nSQL As String
'    Dim nSQLId As String
'    Dim nTitulo As String
'
'    If Col <> 3 Then Exit Sub
'    ReDim xCampos(4, 3) As String
'    xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombres": xCampos(0, 2) = "4500":  xCampos(0, 3) = "C"
'    xCampos(1, 0) = "Peso":         xCampos(1, 1) = "peso":    xCampos(1, 2) = "800":   xCampos(1, 3) = "N"
'    xCampos(2, 0) = "Abrev":        xCampos(2, 1) = "abrev":   xCampos(2, 2) = "700":   xCampos(2, 3) = "C"
'    xCampos(3, 0) = "Id":           xCampos(3, 1) = "id":      xCampos(3, 2) = "500":   xCampos(3, 3) = "N"
'
'    nTitulo = "Buscando Contenedor"
'
'    nSQL = "SELECT pro_pesotara.id, pro_pesotara.descripcion AS nombres, pro_pesotara.peso, pro_pesotara.abrev " _
'        + vbCr + " FROM mae_unidades INNER JOIN pro_pesotara ON mae_unidades.id = pro_pesotara.idundori "
'
'    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombres", "nombres", Principio, ""
'
'    If xRs.State = 0 Then GoTo SALIR
'    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
'
'    Agregando = True
'    '--aplicando filtro
'    RstGrDetTara.Filter = "codigo = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & " and tipo=0 and idemp=" & NulosN(Fg1.TextMatrix(Fg1.Row, 16))
'    If RstGrDetTara.RecordCount <> 0 Then RstGrDetTara.MoveFirst
'    RstGrDetTara.Find "item = " & NulosN(fg(1).TextMatrix(Row, 7))
'    If RstGrDetTara.EOF = False And RstGrDetTara.BOF = False Then
'        RstGrDetTara("abrev") = NulosC(xRs("abrev"))
'        RstGrDetTara("idpeso") = NulosN(xRs("id"))
'        RstGrDetTara("pesouni") = NulosN(xRs("peso"))
'    End If
'    '------------------------------------------------------------------
'    '--actualizar el grid
'    fg(1).TextMatrix(Row, 3) = NulosC(xRs("abrev"))
'    fg(1).TextMatrix(Row, 6) = NulosN(xRs("id")) '--idpeso
'    fg(1).TextMatrix(Row, 4) = NulosN(xRs("peso"))
'
'    Agregando = False
'    fg_CellChanged 1, Row, 1
'    fg(1).SetFocus
'    Set xRs = Nothing
'    Exit Sub
'SALIR:
'    Set xRs = Nothing
'    Agregando = False
'End Sub
'
'Private Sub Fg_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If Index <> 1 Then Exit Sub
'    If KeyCode = vbKeyEscape Then
'        CmdEditor_Click 2
'    ElseIf KeyCode = vbKeyF6 Then
'        Fg1_KeyDown 117, 0
'    End If
'
'End Sub
'
'Private Sub Fg_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'    On Error GoTo error
'    If Index <> 1 Then Exit Sub
'    If KeyCode = 114 Or KeyCode = vbKeyInsert Then CmdEditor_Click 0  'F3 = Agregar Item
'    If KeyCode = 115 Or KeyCode = vbKeyDelete Then CmdEditor_Click 1         'F4 = Eliminar Item
'    Exit Sub
'error:
'    SHOW_ERROR Me.Name, "Fg_KeyUp (" & Index & ")"
'End Sub
'
'Private Sub pRegistroAddTara()
'    '--0 agregar personal; 1 agregar peso-taras
'    Dim mCol%
'    Dim fInsertar As Boolean
'    Agregando = True
'
''    If fg(1).Rows > fg(1).FixedRows Then
''        If NulosN(fg(1).TextMatrix(fg(1).Rows - 1, 1)) = 0 Then '--
''            MsgBox "Ingrese el Peso Bruto", vbExclamation, xTitulo
''        Else
''            fInsertar = True
''        End If
''    Else
'        fInsertar = True
''    End If
'    mCol = 1
'
'    If fInsertar = True Then fg(1).AddItem ""
'
'    fg(1).Row = fg(1).Rows - 1
'    fg(1).Col = mCol
'    '--cargar el buscador por defecto
'    If fInsertar = True Then
'        '---------------------------------------------------------
'        '--agregando el registro
'        mRowAddTara = mRowAddTara + 1
'        RstGrDetTara.AddNew
'        RstGrDetTara("codigo") = NulosN(Fg1.TextMatrix(Fg1.Row, 16))
'        RstGrDetTara("tipo") = 0
'        RstGrDetTara("idemp") = NulosN(Fg1.TextMatrix(Fg1.Row, 16))
'        RstGrDetTara("item") = mRowAddTara
'        fg(1).TextMatrix(fg(1).Rows - 1, 7) = mRowAddTara
'        '---------------------------------------------------------
'        '--colocar el ultimo peso-tara seleccionado
'        If NulosN(RstGrDetTara("idpeso")) = 0 And fg(1).Rows > 2 Then
'            RstGrDetTara("idpeso") = NulosN(fg(1).TextMatrix(fg(1).Rows - 2, 6))
'            RstGrDetTara("pesouni") = NulosN(fg(1).TextMatrix(fg(1).Rows - 2, 4))
'            RstGrDetTara("abrev") = NulosC(fg(1).TextMatrix(fg(1).Rows - 2, 3))
'            RstGrDetTara("cantidad") = NulosN(fg(1).TextMatrix(fg(1).Rows - 2, 2))
'        Else
'            RstGrDetTara("idpeso") = NulosN(FrmControlTarea.lbl_cod(2).Caption)
'            RstGrDetTara("pesouni") = NulosN(FrmControlTarea.lblPesoTara(0).Caption)
'            RstGrDetTara("abrev") = NulosC(FrmControlTarea.lblPesoTara(1).Caption)
'            RstGrDetTara("cantidad") = 1
'        End If
'
'        '---------------------------------------------------------
'        fg(1).TextMatrix(fg(1).Row, 2) = NulosN(RstGrDetTara("cantidad"))
'        fg(1).TextMatrix(fg(1).Row, 3) = NulosC(RstGrDetTara("abrev"))
'        fg(1).TextMatrix(fg(1).Row, 4) = NulosN(RstGrDetTara("pesouni"))
'        fg(1).TextMatrix(fg(1).Row, 6) = NulosN(RstGrDetTara("idpeso"))
'    End If
'    fg(1).SetFocus
'    Agregando = False
'
'End Sub
'
'Private Sub pRegistroDelTara()
'
'    If fg(1).Row < 1 Then
'        MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        fg(1).SetFocus
'        Exit Sub
'    End If
'
'    If fg(1).Rows = 1 Then
'        MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        fg(1).SetFocus
'        Exit Sub
'    End If
'    If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
'    '------------------------------------------------------------------
'    '--aplicando filtro
'    RstGrDetTara.Filter = "codigo = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & " and tipo=0 and idemp=" & NulosN(Fg1.TextMatrix(Fg1.Row, 16))
'    If RstGrDetTara.RecordCount <> 0 Then RstGrDetTara.MoveFirst
'    RstGrDetTara.Find "item = " & NulosN(fg(1).TextMatrix(fg(1).Row, 7))
'    If RstGrDetTara.EOF = False And RstGrDetTara.BOF = False Then
'        RstGrDetTara.Delete
'    End If
'    '------------------------------------------------------------------
'
'    fg(1).RemoveItem fg(1).Row
'    If fg(1).Rows > 1 Then
'        fg(1).Row = fg(1).Rows - 1
'        fg(1).Col = 1
'        fg(1).SetFocus
'    Else
'        CmdEditor(0).SetFocus
'    End If
'
'
'End Sub
'
'
''******************************************************************************
''******************************************************************************
'
'
'Private Sub pCargarTareasReceta()
'    If NulosN(lbl_cod(3).Caption) = 0 Then
'        txt_cb(3).SetFocus
'        Exit Sub
'    End If
'    Dim RstTmp As New ADODB.Recordset
'    Dim nSQL As String
'    On Error GoTo error
'    Me.MousePointer = vbHourglass
'    If optTarea(0).Value = True Then
'        nSQL = "SELECT alm_inventario.descripcion, pro_receta.codrec, pro_receta.id " _
'            + vbCr + " FROM alm_inventario INNER JOIN (pro_tareas INNER JOIN (pro_receta INNER JOIN pro_recetatar ON pro_receta.id = pro_recetatar.idrec) ON pro_tareas.id = pro_recetatar.idtar) ON alm_inventario.id = pro_receta.iditem " _
'            + vbCr + " WHERE (((pro_tareas.id)=" & NulosN(lbl_cod(3).Caption) & "))  " _
'            + vbCr + " ORDER BY alm_inventario.descripcion, pro_receta.codrec;"
'    Else
'        nSQL = "SELECT pro_tareas.id, pro_tareas.descripcion, pro_recetatar.orden " _
'            + vbCr + " FROM pro_tareas INNER JOIN pro_recetatar ON pro_tareas.id = pro_recetatar.idtar " _
'            + vbCr + " Where (((pro_recetatar.idrec) =" & NulosN(lbl_cod(3).Caption) & ")) " _
'            + vbCr + " ORDER BY pro_recetatar.orden;"
'
'    End If
'
'    RST_Busq RstTmp, nSQL, xCon
'    DoEvents
'    Agregando = True
'    fg(0).ColWidth(3) = 0
'
'    If optTarea(0).Value = True Then
'        fg(0).ColWidth(2) = 1020
'        fg(0).TextMatrix(0, 1) = "Producto"
'    Else
'        fg(0).ColWidth(2) = 0
'        fg(0).TextMatrix(0, 1) = "Tarea"
'    End If
'
'    If RstTmp.RecordCount <> 0 Then
'        DoEvents
'        With fg(0)
'            .Rows = 1
'            RstTmp.MoveFirst
'            Do While Not RstTmp.EOF
'                DoEvents
'                .Rows = .Rows + 1
'                .TextMatrix(.Rows - 1, 1) = NulosC(RstTmp.Fields("descripcion"))
'                If optTarea(0).Value = True Then .TextMatrix(.Rows - 1, 2) = NulosC(RstTmp.Fields("codrec"))
'                .TextMatrix(.Rows - 1, 3) = NulosN(RstTmp.Fields("id"))
'                '---
'                RstTmp.MoveNext
'            Loop
'        End With
'    End If
'    Set RstTmp = Nothing
'    Agregando = False
'    Me.MousePointer = vbDefault
'    Exit Sub
'error:
'    Set RstTmp = Nothing
'    Agregando = False
'    Me.MousePointer = vbDefault
'    SHOW_ERROR Me.Name, "pCargarTareasReceta"
'
'End Sub
'
'
'
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub
