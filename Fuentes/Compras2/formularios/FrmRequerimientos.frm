VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmRequerimientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compras - Orden de Requerimiento"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7230
      Left            =   0
      TabIndex        =   8
      Top             =   360
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   12525
         TabIndex        =   9
         Top             =   375
         Width           =   11790
         Begin VB.CommandButton CmdBudSol 
            Height          =   240
            Left            =   2040
            Picture         =   "FrmRequerimientos.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   855
            Width           =   240
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   2040
            Picture         =   "FrmRequerimientos.frx":0132
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   1800
            Width           =   240
         End
         Begin VB.CommandButton CmdBusArea 
            Height          =   240
            Left            =   2040
            Picture         =   "FrmRequerimientos.frx":0264
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   1170
            Width           =   240
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2505
            Locked          =   -1  'True
            TabIndex        =   1
            Text            =   "TxtNumDoc"
            Top             =   510
            Width           =   1365
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   6
            Text            =   "TxtIdMon"
            Top             =   1770
            Width           =   780
         End
         Begin VB.TextBox TxtIdArea 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   3
            Text            =   "TxtIdArea"
            Top             =   1140
            Width           =   780
         End
         Begin VB.TextBox TxtidSol 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   2
            Text            =   "TxtidSol"
            Top             =   825
            Width           =   780
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   0
            Text            =   "TxtNumSer"
            Top             =   510
            Width           =   780
         End
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   4425
            Left            =   75
            TabIndex        =   17
            Top             =   2370
            Width           =   11610
            _cx             =   20479
            _cy             =   7805
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9
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
            BackTabColor    =   -2147483633
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483630
            Caption         =   "        Items        |    Proveedores    | Cotizar Productos "
            Align           =   0
            CurrTab         =   2
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
            Begin VB.Frame Frame9 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   0  'None
               Caption         =   "Frame9"
               Height          =   4065
               Left            =   12225
               TabIndex        =   36
               Top             =   15
               Width           =   11580
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H00FFFFC0&
               BorderStyle     =   0  'None
               Caption         =   "Frame6"
               Height          =   4065
               Left            =   15
               TabIndex        =   20
               Top             =   15
               Width           =   11580
               Begin VSFlex7Ctl.VSFlexGrid Fg3 
                  Height          =   3390
                  Left            =   30
                  TabIndex        =   30
                  Top             =   15
                  Width           =   11520
                  _cx             =   20320
                  _cy             =   5980
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
                  FormatString    =   $"FrmRequerimientos.frx":0396
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
               Begin VB.Frame Frame8 
                  Height          =   690
                  Left            =   30
                  TabIndex        =   31
                  Top             =   3360
                  Width           =   11550
                  Begin VB.CommandButton Command2 
                     Caption         =   "Command2"
                     Height          =   360
                     Left            =   8130
                     TabIndex        =   35
                     Top             =   225
                     Visible         =   0   'False
                     Width           =   960
                  End
                  Begin VB.CommandButton CmdConfCot 
                     Caption         =   "Configurar Cotizacion"
                     Height          =   420
                     Left            =   9105
                     TabIndex        =   32
                     Top             =   180
                     Width           =   1245
                  End
                  Begin VB.Label LblProveedor 
                     BackColor       =   &H00FFFFFF&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "LblProveedor"
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
                     Left            =   1035
                     TabIndex        =   34
                     Top             =   240
                     Width           =   6945
                  End
                  Begin VB.Label Label10 
                     Caption         =   "Proveedor"
                     Height          =   210
                     Left            =   135
                     TabIndex        =   33
                     Top             =   285
                     Width           =   900
                  End
               End
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   4065
               Left            =   -12195
               TabIndex        =   19
               Top             =   15
               Width           =   11580
               Begin VSFlex7Ctl.VSFlexGrid Fg2 
                  Height          =   3390
                  Left            =   30
                  TabIndex        =   25
                  Top             =   15
                  Width           =   11520
                  _cx             =   20320
                  _cy             =   5980
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
                  FormatString    =   $"FrmRequerimientos.frx":03FC
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
                  Height          =   690
                  Left            =   30
                  TabIndex        =   26
                  Top             =   3360
                  Width           =   11550
                  Begin VB.CommandButton Command4 
                     Caption         =   "Agregar Nuevo Proveedor"
                     Height          =   420
                     Left            =   3225
                     TabIndex        =   29
                     Top             =   180
                     Visible         =   0   'False
                     Width           =   1800
                  End
                  Begin VB.CommandButton CmdDelPro 
                     Caption         =   "Eliminar Proveedor"
                     Height          =   420
                     Left            =   1590
                     TabIndex        =   28
                     Top             =   180
                     Width           =   1245
                  End
                  Begin VB.CommandButton CmdAddPro 
                     Caption         =   "Agregar Proveedor"
                     Height          =   420
                     Left            =   315
                     TabIndex        =   27
                     Top             =   180
                     Width           =   1245
                  End
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame4"
               Height          =   4065
               Left            =   -12495
               TabIndex        =   18
               Top             =   15
               Width           =   11580
               Begin VSFlex7Ctl.VSFlexGrid Fg1 
                  Height          =   3390
                  Left            =   30
                  TabIndex        =   7
                  Top             =   15
                  Width           =   11520
                  _cx             =   20320
                  _cy             =   5980
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
                  Cols            =   10
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmRequerimientos.frx":047F
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
               Begin VB.Frame Frame3 
                  Height          =   690
                  Left            =   30
                  TabIndex        =   21
                  Top             =   3360
                  Width           =   11550
                  Begin VB.CommandButton CmdAddItem 
                     Caption         =   "Agregar Item"
                     Height          =   420
                     Left            =   315
                     TabIndex        =   24
                     Top             =   180
                     Width           =   1245
                  End
                  Begin VB.CommandButton CmdDelItem 
                     Caption         =   "Eliminar Item"
                     Height          =   420
                     Left            =   1590
                     TabIndex        =   23
                     Top             =   180
                     Width           =   1245
                  End
                  Begin VB.CommandButton Command1 
                     Caption         =   "Agregar Nuevo Item"
                     Height          =   420
                     Left            =   3225
                     TabIndex        =   22
                     Top             =   180
                     Width           =   1800
                  End
               End
            End
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEnt 
            Height          =   300
            Left            =   5145
            TabIndex        =   5
            Top             =   1470
            Width           =   1200
            _ExtentX        =   2117
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
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmi 
            Height          =   300
            Left            =   1530
            TabIndex        =   4
            Top             =   1455
            Width           =   1200
            _ExtentX        =   2117
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
         End
         Begin VB.Label LblSolicitante 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblSolicitante"
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
            TabIndex        =   48
            Top             =   825
            Width           =   5490
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
            Left            =   2370
            TabIndex        =   47
            Top             =   1770
            Width           =   2280
         End
         Begin VB.Label LblArea 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblArea"
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
            TabIndex        =   46
            Top             =   1140
            Width           =   2280
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   1830
            Width           =   585
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Entrega"
            Height          =   195
            Left            =   4155
            TabIndex        =   44
            Top             =   1515
            Width           =   915
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Emisión"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   1515
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Area"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   1200
            Width           =   330
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Solicitante"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   870
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nº Requerimiento"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   570
            Width           =   1245
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "[  Lista de Items  ]"
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
            Height          =   195
            Left            =   135
            TabIndex        =   16
            Top             =   2145
            Width           =   1560
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Orden de Requerimiento"
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
            TabIndex        =   10
            Top             =   30
            Width           =   11610
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   45
         TabIndex        =   11
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6480
            Left            =   30
            TabIndex        =   12
            Top             =   300
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11430
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "id"
            Columns(0).DataField=   "id"
            Columns(0).NumberFormat=   "General Number"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Documento"
            Columns(1).DataField=   "descdoc"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nº Documento"
            Columns(2).DataField=   "numdoc2"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch. Emi."
            Columns(3).DataField=   "fchemi"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fch. Ent."
            Columns(4).DataField=   "fchent"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Area"
            Columns(5).DataField=   "descarea"
            Columns(5).NumberFormat=   "0.00"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Solicitante"
            Columns(6).DataField=   "apenom"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "M"
            Columns(7).DataField=   "descmon"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   16
            Columns(8)._MaxComboItems=   5
            Columns(8).ValueItems(0)._DefaultItem=   0
            Columns(8).ValueItems(0).Value=   "1"
            Columns(8).ValueItems(0).Value.vt=   8
            Columns(8).ValueItems(0).DisplayValue=   "Pendiente"
            Columns(8).ValueItems(0).DisplayValue.vt=   8
            Columns(8).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(8).ValueItems(1)._DefaultItem=   0
            Columns(8).ValueItems(1).Value=   "2"
            Columns(8).ValueItems(1).Value.vt=   8
            Columns(8).ValueItems(1).DisplayValue=   "Cotizada"
            Columns(8).ValueItems(1).DisplayValue.vt=   8
            Columns(8).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(8).ValueItems(2)._DefaultItem=   0
            Columns(8).ValueItems(2).Value=   "3"
            Columns(8).ValueItems(2).Value.vt=   8
            Columns(8).ValueItems(2).DisplayValue=   "Orden de Compra"
            Columns(8).ValueItems(2).DisplayValue.vt=   8
            Columns(8).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
            Columns(8).ValueItems.Count=   3
            Columns(8).Caption=   "Estado"
            Columns(8).DataField=   "idest"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   9
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   397
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=9"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=873"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=794"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=3096"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3016"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=2487"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2408"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1799"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1720"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1852"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1773"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=3440"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=3360"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=4313"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=4233"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=512"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=767"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=688"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=512"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(50)=   "Column(8).Width=1720"
            Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=1640"
            Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=516"
            Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=78,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=75,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=76,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=77,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=86,.parent=13,.alignment=0"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=83,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=84,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=85,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=82,.parent=13,.alignment=0"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=79,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=80,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=81,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=28,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=25,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=26,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=27,.parent=17"
            _StyleDefs(72)  =   "Named:id=33:Normal"
            _StyleDefs(73)  =   ":id=33,.parent=0"
            _StyleDefs(74)  =   "Named:id=34:Heading"
            _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(76)  =   ":id=34,.wraptext=-1"
            _StyleDefs(77)  =   "Named:id=35:Footing"
            _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(79)  =   "Named:id=36:Selected"
            _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(81)  =   "Named:id=37:Caption"
            _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(83)  =   "Named:id=38:HighlightRow"
            _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(85)  =   "Named:id=39:EvenRow"
            _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(87)  =   "Named:id=40:OddRow"
            _StyleDefs(88)  =   ":id=40,.parent=33"
            _StyleDefs(89)  =   "Named:id=41:RecordSelector"
            _StyleDefs(90)  =   ":id=41,.parent=34"
            _StyleDefs(91)  =   "Named:id=42:FilterBar"
            _StyleDefs(92)  =   ":id=42,.parent=33"
         End
         Begin VB.Label LblMes 
            AutoSize        =   -1  'True
            Caption         =   "LblMes"
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
            TabIndex        =   15
            Top             =   30
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Requerimientos"
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
            TabIndex        =   14
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label LblPeriodo 
            Alignment       =   2  'Center
            Caption         =   "LblPeriodo"
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
            TabIndex        =   13
            Top             =   0
            Width           =   1860
         End
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
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRequerimientos.frx":05AC
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRequerimientos.frx":0AF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRequerimientos.frx":0E82
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRequerimientos.frx":1006
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRequerimientos.frx":145A
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRequerimientos.frx":1572
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRequerimientos.frx":1AB6
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRequerimientos.frx":1FFA
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRequerimientos.frx":210E
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRequerimientos.frx":2222
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRequerimientos.frx":2676
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRequerimientos.frx":27E2
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRequerimientos.frx":2D2A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir "
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
End
Attribute VB_Name = "FrmRequerimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstLista As New ADODB.Recordset
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim CaracteresNumericos As String

Public DeDonde As Integer                ' ESPECIFICA DESDE DONDE SE ESTA LLAMANDO AL FORMULARIO
                                         ' 1 = MENU DEL SISTEMAS
                                         ' 2 = OTRO FORMULARIO
Public xIdOR As Integer                     ' ESPECIFICA EL ID DE LA ORDEN QUE SE MOSTRARA CUANDO LA VARIABLE DEDONDE SEA 2
Dim fOrdenLista As Boolean                 ' --especfica el orden de la lista de la consulta
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


Private Sub CmdAddItem_Click()
    If QueHace = 3 Then Exit Sub
    If NulosC(Fg1.TextMatrix(Fg1.Rows - 1, 2)) = "" Then
        Exit Sub
    End If
    Fg1.Rows = Fg1.Rows + 1
End Sub

Private Sub CmdAddPro_Click()
    If QueHace = 3 Then Exit Sub
    If NulosC(Fg2.TextMatrix(Fg2.Rows - 1, 2)) = "" Then
        Exit Sub
    End If
    Fg2.Rows = Fg2.Rows + 1
End Sub

Private Sub CmdBudSol_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Empleado":    xCampos(0, 1) = "apenom":           xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":      xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xForm.SQLCad = "SELECT com_usuario.id, com_usuario.idper, UCase([pla_empleados]![apepat]) & ' ' & UCase([pla_empleados]![apemat]) & ', ' & [pla_empleados]![nom] AS apenom " _
        & " FROM com_usuario LEFT JOIN pla_empleados ON com_usuario.idper = pla_empleados.id"

    xForm.Titulo = "Buscando Usuarios"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "apenom"
    xForm.CampoBusca = "apenom"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtidSol.Text = xRs("id")
            LblSolicitante.Caption = xRs("apenom")
            TxtIdArea.SetFocus
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusArea_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xForm.SQLCad = "SELECT * FROM mae_area ORDER BY descripcion"
    
    xForm.Titulo = "Buscando Area"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "descripcion"
    xForm.CampoBusca = "descripcion"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdArea.Text = xRs("id")
            LblArea.Caption = xRs("descripcion")
            TxtFchEmi.SetFocus
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMon_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xForm.SQLCad = "SELECT * FROM mae_moneda ORDER BY descripcion"
    
    xForm.Titulo = "Buscando Area"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "descripcion"
    xForm.CampoBusca = "descripcion"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdMon.Text = xRs("id")
            LblMoneda.Caption = xRs("descripcion")
            Fg1.SetFocus
        End If
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdConfCot_Click()
    'If QueHace = 3 Then Exit Sub
    
    Dim A As Integer
    Dim B As Integer
    
    ' ELIMINAMOS LAS FILAS EN BLANCO DEL CONTROL Fg1
    For A = 1 To Fg1.Rows - 1
        If Fg1.TextMatrix(A, 2) = "" Then
            Fg1.RemoveItem A
        End If
    Next A
    
    ' ELIMINAMOS LAS FILAS EN BLANCO DEL CONTROL Fg2
    For A = 1 To Fg2.Rows - 1
        If Fg2.TextMatrix(A, 1) = "" Then
            Fg2.RemoveItem A
        End If
    Next A
    
    Fg3.Rows = 1
    Fg3.Cols = 3
    For B = 1 To Fg2.Rows - 1
        Fg3.Cols = Fg3.Cols + 1
        Fg3.TextMatrix(0, Fg3.Cols - 1) = Fg2.TextMatrix(B, 2)
        Fg3.ColDataType(Fg3.Cols - 1) = flexDTBoolean
    Next B
    
    For A = 1 To Fg1.Rows - 1
        Fg3.Rows = Fg3.Rows + 1
        Fg3.TextMatrix(Fg3.Rows - 1, 1) = Fg1.TextMatrix(A, 2)
        Fg3.TextMatrix(Fg3.Rows - 1, 2) = Fg1.TextMatrix(A, 7)
    Next A
End Sub

Private Sub CmdDelItem_Click()
    If QueHace = 3 Then Exit Sub
    If Fg1.Rows = 1 Then
        MsgBox "No ha items para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    Fg1.RemoveItem Fg1.Row
    Dim A As Integer
    
    For A = 1 To Fg1.Rows - 1
        Fg1.TextMatrix(A, 1) = Str(A)
    Next A
End Sub

Private Sub CmdDelPro_Click()
    If QueHace = 3 Then Exit Sub
    If Fg2.Rows = 1 Then
        MsgBox "No ha items para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    Fg2.RemoveItem Fg2.Row
    Dim A As Integer
    
    For A = 1 To Fg2.Rows - 1
        Fg2.TextMatrix(A, 1) = Str(A)
    Next A
End Sub

Private Sub Command1_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xFun As New Sgi2_Procesos.Procesos
    Dim xIdProducto As Integer
    Dim xRs As New ADODB.Recordset
    
    xIdProducto = xFun.IngRapidoItems(xCon)
    If xIdProducto <> 0 Then
        RST_Busq xRs, "SELECT alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, mae_tipoproducto.descripcion AS desctippro, alm_inventario.id, " _
            & " alm_inventario.idunimed FROM mae_tipoproducto RIGHT JOIN (mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed) " _
            & " ON mae_tipoproducto.id = alm_inventario.tippro Where (((alm_inventario.activo) = -1) And ((alm_inventario.id) = " & xIdProducto & ")) ORDER BY alm_inventario.descripcion", xCon
    
        If xRs.RecordCount <> 0 Then
            Fg1.TextMatrix(Fg1.Row, 2) = xRs("descripcion")
            Fg1.TextMatrix(Fg1.Row, 3) = xRs("abrev")
            Fg1.TextMatrix(Fg1.Row, 7) = xRs("id")
            Fg1.TextMatrix(Fg1.Row, 8) = xRs("idunimed")
        End If
        Set xRs = Nothing
        Fg1.SetFocus
    End If
End Sub

Private Sub Command2_Click()
'    Dim Li As Integer
'    Dim strSource As String
'    Dim xArea, xEmp, xDir, xCuerpo As String
'    Dim xEmpleado As String
'    Dim Pagina As Integer
'    Dim Lineas As Integer
'
'    'On Error Resume Next
'    Set oPDF = New cPDF
'
'    If oPDF.PDFCreate(App.Path & "\holas.pdf") = True Then
'        oPDF.Fonts.Add "Tit", Times_BoldItalic, WinAnsiEncoding
'        oPDF.Fonts.Add "Head", Times_Italic, WinAnsiEncoding
'        oPDF.Fonts.Add "Cont", Courier, WinAnsiEncoding
'        oPDF.Fonts.Add "Time", Times_Roman, WinAnsiEncoding
'        CrearCabecera
'
'        oPDF.WTextBox 100, 55, 10, 420, "Villa el Salvador 25 de Noviembre del 2009", "Time", 9, hJustify, vMiddle, vbBlack, 1, vbBlack
'        oPDF.WTextBox 120, 55, 10, 420, "Para :", "Time", 9, hJustify, vMiddle, vbBlack, 1, vbBlack
'
'        xArea = "Departamento de Ventas"
'        xEmp = "DISTRIBUIDORA DEL SUR S.A.C."
'        xDir = "Av. Revolución Industrial 158 Cercado de Lima"
'
'        oPDF.WTextBox 130, 80, 10, 420, xArea, "Time", 9, hJustify, vMiddle, vbBlack, 1, vbBlack
'        oPDF.WTextBox 140, 80, 10, 420, xEmp, "Time", 9, hJustify, vMiddle, vbBlack, 1, vbBlack
'        oPDF.WTextBox 150, 80, 10, 420, xDir, "Time", 9, hJustify, vMiddle, vbBlack, 1, vbBlack
'
'        ' ESCRIBIMOS EL CONTENIDO DEL CUERPO
'        xCuerpo = "Por medio de la presente le saludamos y solicitamos nos envié en el mas breve plazo la cotización de los siguientes ítems"
'        oPDF.WTextBox 170, 55, 10, 420, xCuerpo, "Time", 9, hJustify, vMiddle, vbBlack, 1, vbBlack
'
'        ' ESCRIBIMOS EL FINAL DEL DOCUMENTO
'        xCuerpo = "Especificar:"
'        oPDF.WTextBox 470, 55, 10, 420, xCuerpo, "Time", 9, hJustify, vMiddle, vbBlack, 1, vbBlack
'        xCuerpo = "* Validez de la oferta "
'        oPDF.WTextBox 480, 75, 10, 420, xCuerpo, "Time", 9, hJustify, vMiddle, vbBlack, 1, vbBlack
'        xCuerpo = "* Lugar de entrega"
'        oPDF.WTextBox 490, 75, 10, 420, xCuerpo, "Time", 9, hJustify, vMiddle, vbBlack, 1, vbBlack
'        xCuerpo = "* Formato de pago"
'        oPDF.WTextBox 500, 75, 10, 420, xCuerpo, "Time", 9, hJustify, vMiddle, vbBlack, 1, vbBlack
'        xCuerpo = "* Plazo de entrega"
'        oPDF.WTextBox 510, 75, 10, 420, xCuerpo, "Time", 9, hJustify, vMiddle, vbBlack, 1, vbBlack
'
'        xCuerpo = "Atentamente"
'        oPDF.WTextBox 530, 55, 10, 420, xCuerpo, "Time", 9, hCenter, vMiddle, vbBlack, 1, vbBlack
'
'        ' ESCRIBIMOS LA FIRMA DEL ENCARGADO
'        xCuerpo = "--------------------------------"
'        oPDF.WTextBox 560, 55, 10, 420, xCuerpo, "Time", 9, hCenter, vMiddle, vbBlack, 1, vbBlack
'        xEmpleado = "Juan Perez Martinez"
'        oPDF.WTextBox 570, 55, 10, 420, xEmpleado, "Time", 9, hCenter, vMiddle, vbBlack, 1, vbBlack
'        xCuerpo = "Jefe de Compras"
'        oPDF.WTextBox 580, 55, 10, 420, xCuerpo, "Time", 9, hCenter, vMiddle, vbBlack, 1, vbBlack
'        Li = Li + 15
'
'        oPDF.WLineTo 493, Li, 32, Li
'        oPDF.LineStroke
'        Li = Li + 5
'        oPDF.PDFClose
'        Set oPDF = Nothing
'    Else
'        MsgBox "No se Puede Mostrar Documento", vbCritical, "Error"
'    End If
End Sub

'Sub CrearCabecera()
'    Dim xNomEmp, xDirEmp, xTelEmp, xMaiEmp As String
'    Dim xNumDoc As String
'    xNomEmp = "AGRO INDUSTRIAS EL VADO E.I.R.L."
'    xDirEmp = "Mz K1 Lte 11 Parcela II Parque Industrial V.E.S."
'    xTelEmp = "Telf: 493-0808   Tele Fax: 295-6868"
'    xMaiEmp = "Pag. WEB  www.agro-vado.com"
'    xNumDoc = "0001-00000001"
'
'    Dim Pagina As Integer
'    oPDF.NewPage UsarAnchoAlto, 525, 675
'    Pagina = Pagina + 1
'    oPDF.WTextBox 32, 55, 20, 250, xNomEmp, "Tit", 12, hCenter, vMiddle, RGB(0, 0, 128), 2, vbRed
'    oPDF.WTextBox 55, 55, 10, 250, xDirEmp, "Head", 9, hCenter, vMiddle, RGB(0, 0, 128), 2, vbRed
'    oPDF.WTextBox 65, 55, 10, 250, xTelEmp, "Head", 9, hCenter, vMiddle, RGB(0, 0, 128), 2, vbRed
'    oPDF.WTextBox 75, 55, 10, 250, xMaiEmp, "Head", 9, hCenter, vMiddle, RGB(0, 0, 128), 2, vbRed
'
'    oPDF.WTextBox 46, 330, 10, 150, "ORDEN DE COTIZACION", "Head", 10, hCenter, vMiddle, RGB(0, 0, 128), 2, vbRed
'    oPDF.WTextBox 60, 330, 10, 150, "Nº " & xNumDoc, "Head", 10, hCenter, vMiddle, RGB(0, 0, 128), 2, vbRed
'
'    oPDF.WRectangle 32, 330, 53, 150, 1.5, vbBlack
'End Sub

'Function CrearHeader1() As Boolean
'    Dim Pagina As Integer
'    oPDF.NewPage UsarAnchoAlto, 525, 675
'    Pagina = Pagina + 1
'    oPDF.WTextBox 32, 32, 20, 461, "Reporte de Productos Ordenados por Bodega", "Tit", 20, hCenter, vMiddle, RGB(0, 0, 128), 0, vbRed
'    oPDF.WTextBox 55, 32, 10, 461, "Nit : " & "1010101010", "Head", 9, hCenter, vMiddle, RGB(0, 0, 128), 0, vbRed
'    oPDF.WTextBox 65, 32, 10, 461, "EMPRESA :" & " : " & "INDUSTRIAS EL VADO", "Head", 9, hCenter, vMiddle, RGB(0, 0, 128), 0, vbRed
'    oPDF.WTextBox 65, 32, 10, 461, "Fecha: " & CStr(Format$(CDate(Now), "dd \de MMMM \de yyyy")), "Head", 9, hRight, vMiddle, RGB(0, 0, 128), 0, vbRed
'
'    oPDF.WLineTo 493, 78, 32, 78
'    oPDF.WLineTo 493, 79, 32, 79
'    oPDF.LineStroke
'
'    oPDF.WTextBox 85, 32, 10, 60, "Codigo", "Cont", 9, hCenter, vMiddle, vbBlack, 0, vbBlack
'    oPDF.WTextBox 85, 92, 10, 200, "Nombre", "Cont", 9, hCenter, vMiddle, vbBlack, 0, vbBlack
'    oPDF.WTextBox 85, 292, 10, 81, "Lote", "Cont", 9, hCenter, vMiddle, vbBlack, 0, vbBlack
'    oPDF.WTextBox 85, 373, 10, 120, "Fecha", "Cont", 9, hCenter, vMiddle, vbBlack, 0, vbBlack
'
'    oPDF.WLineTo 493, 96, 32, 96
'    oPDF.LineStroke
'End Function

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstLista
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA LAS COLUMNAS DEL CONTROL Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstLista.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstLista("id")), xCon
    End If
End Sub


Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset

    If Col = 2 Then
        Dim xCampos(3, 4) As String
        
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "codpro":           xCampos(1, 2) = "1400":         xCampos(1, 3) = "c"
        xCampos(2, 0) = "Abreviatura":    xCampos(2, 1) = "abrev":            xCampos(2, 2) = "1000":         xCampos(2, 3) = "c"
        xCampos(3, 0) = "Tipo Producto":  xCampos(3, 1) = "desctippro":       xCampos(3, 2) = "1200":         xCampos(3, 3) = "c"
        
        xForm.SQLCad = "SELECT alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, mae_tipoproducto.descripcion AS desctippro, " _
            & " alm_inventario.id, alm_inventario.idunimed FROM mae_tipoproducto RIGHT JOIN (mae_unidades RIGHT JOIN alm_inventario " _
            & " ON mae_unidades.id = alm_inventario.idunimed) ON mae_tipoproducto.id = alm_inventario.tippro Where (((alm_inventario.activo) = -1)) " _
            & " ORDER BY alm_inventario.descripcion"
        
        xForm.Titulo = "Buscando Items"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 2) = xRs("descripcion")
                Fg1.TextMatrix(Fg1.Row, 3) = xRs("abrev")
                Fg1.TextMatrix(Fg1.Row, 7) = xRs("id")
                Fg1.TextMatrix(Fg1.Row, 8) = xRs("idunimed")
                Fg1.TextMatrix(Fg1.Row, 1) = (NulosN(Fg1.TextMatrix(Fg1.Row - 1, 1)) + 1)
                
                If NulosC(Fg1.TextMatrix(Fg1.Rows - 1, 2)) <> "" Then
                    Fg1.Rows = Fg1.Rows + 1
                    Fg1.TextMatrix(Fg1.Rows - 1, 1) = (NulosN(Fg1.TextMatrix(Fg1.Rows - 2, 1)) + 1)
                End If
            End If
        End If
    End If
    
    If Col = 5 Then
        If NulosN(TxtIdArea.Text) = 0 Then
            MsgBox "No ha especificado el area que solicita la orden de compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtIdArea.SetFocus
            Exit Sub
        End If
        
        Dim xCampos2(2, 4) As String
        
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos2(0, 0) = "Codigo":        xCampos2(0, 1) = "codigo":          xCampos2(0, 2) = "1200":         xCampos2(0, 3) = "C"
        xCampos2(1, 0) = "Descripcion":   xCampos2(1, 1) = "descripcion":     xCampos2(1, 2) = "6000":         xCampos2(1, 3) = "C"
        
        xForm.SQLCad = "SELECT con_centrocosto.* FROM con_centocostoarea LEFT JOIN con_centrocosto ON con_centocostoarea.idcencos = con_centrocosto.id " _
            & " Where (((con_centocostoarea.idarea) = " & NulosN(TxtIdArea.Text) & ")) ORDER BY con_centrocosto.codigo"
        
        xForm.Titulo = "Buscando Centros de Costos"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "codigo"
        xForm.CampoBusca = "codigo"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos2)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 5) = Trim(xRs("codigo"))
                Fg1.TextMatrix(Fg1.Row, 6) = Trim(xRs("descripcion"))
                Fg1.TextMatrix(Fg1.Row, 9) = xRs("id")
            End If
        End If
    End If
    
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then Exit Sub
    If Fg1.Col = 2 Or Fg1.Col = 4 Or Fg1.Col = 5 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = 4 Or Col = 5 Then
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    If Col = 1 Then
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "nombre":           xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Nº R.U.C.":      xCampos(1, 1) = "numruc":           xCampos(1, 2) = "1400":         xCampos(1, 3) = "c"
        
        xForm.SQLCad = "SELECT mae_prov.numruc, mae_prov.nombre, mae_prov.id From mae_prov ORDER BY mae_prov.nombre"
        
        xForm.Titulo = "Buscando Proveedores"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "nombre"
        xForm.CampoBusca = "nombre"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg2.TextMatrix(Fg2.Row, 1) = xRs("numruc")
                Fg2.TextMatrix(Fg2.Row, 2) = xRs("nombre")
                Fg2.TextMatrix(Fg2.Row, 3) = xRs("id")
            
                If NulosC(Fg2.TextMatrix(Fg2.Rows - 1, 2)) <> "" Then
                    Fg2.Rows = Fg2.Rows + 1
                End If
            End If
        End If
    End If
End Sub

Private Sub Fg3_EnterCell()
    If Fg3.Col <= 2 Then
        LblProveedor.Caption = ""
        Exit Sub
    End If
    
    LblProveedor.Caption = Fg3.TextMatrix(0, Fg3.Col)
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        CargarLista
        If RstLista.RecordCount = 0 Then
            Dim Rpta As Integer
            
            Rpta = MsgBox("No se ha registrado requerimiento ¿Desea agregar uno ahora?", vbYesNo + vbQuestion + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                Nuevo
            Else
                Set RstLista = Nothing
                Unload Me
            End If
        End If
        If DeDonde = 2 Then
            Toolbar1.Buttons(1).Visible = False
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(3).Visible = False
            Toolbar1.Buttons(4).Visible = False
            Toolbar1.Buttons(5).Visible = False
            Toolbar1.Buttons(6).Visible = False
            Toolbar1.Buttons(7).Visible = False
            TabOne1.CurrTab = 1
        End If
        
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    QueHace = 3
    TabOne1.CurrTab = 0
    CaracteresNumericos = "0123456789." & Chr(8)
    Fg1.ColWidth(7) = 0
    Fg1.ColWidth(8) = 0
    Fg1.ColWidth(9) = 0
    Fg2.ColWidth(3) = 0
    Fg3.ColWidth(2) = 0
    
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Frame4.BackColor = &H8000000F
    Frame5.BackColor = &H8000000F
    Frame6.BackColor = &H8000000F
    Fg1.SelectionMode = flexSelectionByRow
End Sub

Sub Blanquea()
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtidSol.Text = ""
    TxtIdArea.Text = ""
    TxtFchEmi.Valor = ""
    TxtFchEnt.Valor = ""
    TxtIdMon.Text = ""
    LblSolicitante.Caption = ""
    LblArea.Caption = ""
    LblMoneda.Caption = ""
    LblProveedor.Caption = ""
End Sub

Sub Bloquea()
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtidSol.Locked = Not TxtidSol.Locked
    TxtIdArea.Locked = Not TxtIdArea.Locked
    TxtFchEmi.Locked = Not TxtFchEmi.Locked
    TxtFchEnt.Locked = Not TxtFchEnt.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
End Sub

Sub ActivaTool()
    Dim A As Integer
    For A = 1 To 15
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Cancelar()
    Label5.Caption = "Detalle de la Orden de Requerimiento"
    QueHace = 3
    ActivaTool
    Bloquea
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.Editable = flexEDNone
    Dg1.SetFocus
End Sub

Sub Modificar()
    Label5.Caption = "Modificando Orden de Requerimiento"
    QueHace = 2
    xHorIni = Time
    ActivaTool
    TabOne1.TabEnabled(0) = False
    TabOne1.CurrTab = 1
    Bloquea
    Blanquea
    MuestraSegundoTab
    Fg1.SelectionMode = flexSelectionFree
    Fg2.SelectionMode = flexSelectionFree
    Fg3.SelectionMode = flexSelectionFree
    
    Fg1.Editable = flexEDKbdMouse
    Fg2.Editable = flexEDKbdMouse
    Fg3.Editable = flexEDKbdMouse
    
    Fg1.ColComboList(2) = "|..."
    Fg1.ColComboList(5) = "|..."
    
    Fg2.ColComboList(1) = "|..."
    Fg1.Rows = Fg1.Rows + 1
    Fg2.Rows = Fg2.Rows + 1
    TxtidSol.SetFocus
End Sub

Sub Nuevo()
    Label5.Caption = "Agregando Orden de Requerimiento"
    QueHace = 1
    xHorIni = Time
    ActivaTool
    TabOne1.TabEnabled(0) = False
    TabOne1.CurrTab = 1
    TabOne2.CurrTab = 0
    Bloquea
    Blanquea
    TxtNumSer.Text = "0001"
    TxtNumDoc.Text = HallaNumReq(TxtNumSer.Text)
    Fg1.Rows = 1
    Fg1.Rows = Fg1.Rows + 1
    Fg1.ColComboList(2) = "|..."
    Fg1.ColComboList(5) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    Fg1.TextMatrix(Fg1.Rows - 1, 1) = 1
    
    Fg2.Rows = 1
    Fg2.Rows = Fg2.Rows + 1
    Fg2.ColComboList(1) = "|..."
    Fg2.SelectionMode = flexSelectionFree
    
    Fg3.Rows = 1
    Fg3.SelectionMode = flexSelectionFree
    
    Fg1.Editable = flexEDKbdMouse
    Fg2.Editable = flexEDKbdMouse
    Fg3.Editable = flexEDKbdMouse
    TxtidSol.SetFocus
End Sub

Function HallaNumReq(NumeroSerie As String) As String
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT * FROM com_ordenreq WHERE numser = '" & NumeroSerie & "' ORDER BY numdoc", xCon
    If Rst.RecordCount = 0 Then
        HallaNumReq = "0000000001"
    Else
        Rst.MoveLast
        HallaNumReq = Format(Rst("numdoc") + 1, "0000000000")
    End If
    Set Rst = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o editando un registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If RstLista.State = 0 Then Exit Sub
        If RstLista.RecordCount = 0 And QueHace <> 1 Then
            Cancel = 1
            Exit Sub
        End If
        If QueHace = 3 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstLista.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstLista.Filter = ""
    End If
    
    If Button.Index = 15 Then
        Set RstLista = Nothing
        Unload Me
    End If
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    
    Rpta = MsgBox("¿Esta seguro de eliminar el registro seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM com_ordenreq WHERE id = " & RstLista("id") & ""
        xCon.Execute "DELETE * FROM com_ordenreqdet WHERE idor = " & RstLista("id") & ""
        xCon.Execute "DELETE * FROM com_ordenreqpro WHERE idor = " & RstLista("id") & ""
        xCon.Execute "DELETE * FROM com_ordenreqprocot WHERE idor = " & RstLista("id") & ""
        
                'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstLista("id") & " AND idform = " & IdMenuActivo

        
        RstLista.Requery
        Dg1.Refresh
        MsgBox "El registro se elimino con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        

    End If
End Sub

Sub CargarLista()
    
    TDB_FiltroLimpiar Dg1
    Set RstLista = Nothing
    '------------------------------------------

    If DeDonde <> 2 Then
        RST_Busq RstLista, "SELECT com_ordenreq.*, mae_area.descripcion AS descarea, UCase([pla_empleados]![apepat]) & ' ' & UCase([pla_empleados]![apemat]) & ', ' & [pla_empleados]![nom] AS apenom, " _
            & " mae_moneda.simbolo AS descmon, [com_ordenreq]![numser] & '-' & [com_ordenreq]![numdoc] AS numdoc2, mae_documento.descripcion AS descdoc " _
            & " FROM ((((com_ordenreq LEFT JOIN com_usuario ON com_ordenreq.idsol = com_usuario.id) LEFT JOIN mae_area ON com_ordenreq.idarea = mae_area.id) " _
            & " LEFT JOIN mae_moneda ON com_ordenreq.idmon = mae_moneda.id) LEFT JOIN pla_empleados ON com_usuario.idper = pla_empleados.id) LEFT JOIN mae_documento " _
            & " ON com_ordenreq.idtipdoc = mae_documento.id ", xCon
    Else
        RST_Busq RstLista, "SELECT com_ordenreq.*, mae_area.descripcion AS descarea, UCase([pla_empleados]![apepat]) & ' ' & UCase([pla_empleados]![apemat]) & ', ' & [pla_empleados]![nom] AS apenom, " _
            & " mae_moneda.simbolo AS descmon, [com_ordenreq]![numser] & '-' & [com_ordenreq]![numdoc] AS numdoc2, mae_documento.descripcion AS descdoc " _
            & " FROM ((((com_ordenreq LEFT JOIN com_usuario ON com_ordenreq.idsol = com_usuario.id) LEFT JOIN mae_area ON com_ordenreq.idarea = mae_area.id) " _
            & " LEFT JOIN mae_moneda ON com_ordenreq.idmon = mae_moneda.id) LEFT JOIN pla_empleados ON com_usuario.idper = pla_empleados.id) LEFT JOIN mae_documento " _
            & " ON com_ordenreq.idtipdoc = mae_documento.id WHERE (((com_ordenreq.id)=" & xIdOR & "))", xCon
    End If
    Set Dg1.DataSource = RstLista
End Sub

Private Sub TxtIdArea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdArea_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusArea_Click
    End If
End Sub

Private Sub TxtIdArea_Validate(Cancel As Boolean)
    If NulosN(TxtIdArea.Text) = 0 Then
        LblArea.Caption = ""
        Exit Sub
    End If
    
    LblArea.Caption = Busca_Codigo(TxtIdArea.Text, "id", "descripcion", "mae_area", "N", xCon)
    If NulosC(LblArea.Caption) = "" Then
        TxtIdArea.Text = ""
        LblArea.Caption = ""
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
    If NulosN(TxtIdMon.Text) = 0 Then
        LblMoneda.Caption = ""
        Exit Sub
    End If
    
    LblMoneda.Caption = Busca_Codigo(TxtIdMon.Text, "id", "descripcion", "mae_moneda", "N", xCon)
    If NulosC(LblMoneda.Caption) = "" Then
        TxtIdMon.Text = ""
        LblMoneda.Caption = ""
    End If
End Sub

Private Sub TxtIdSol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdSol_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBudSol_Click
    End If
End Sub

Private Sub TxtIdSol_Validate(Cancel As Boolean)
    If NulosN(TxtidSol.Text) = 0 Then
        LblSolicitante.Caption = ""
        Exit Sub
    End If
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT com_usuario.id, com_usuario.idper, UCase([pla_empleados]![apepat]) & ' ' & UCase([pla_empleados]![apemat]) & ', ' & [pla_empleados]![nom] AS apenom " _
        & " FROM com_usuario LEFT JOIN pla_empleados ON com_usuario.idper = pla_empleados.id WHERE (((com_usuario.id)=" & NulosN(TxtidSol.Text) & "))", xCon

    LblSolicitante.Caption = Rst("apenom")
    If NulosC(LblSolicitante.Caption) = "" Then
        TxtidSol.Text = ""
        LblSolicitante.Caption = ""
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
    If NulosC(TxtNumDoc.Text) <> "" Then
        TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
    End If
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Function Grabar() As Boolean
    Dim xCampos(10, 5) As String
    Dim xCampos2(5, 5) As String
    Dim xCampos3(1, 5) As String
    Dim xCampos4(2, 5) As String
    Dim xId As Double
    Dim A, B As Integer
    
    ' ELIMINAMOS LAS FILAS EN BLANCO DEL CONTROL Fg1
    For A = 1 To Fg1.Rows - 1
        If Fg1.TextMatrix(A, 2) = "" Then
            Fg1.RemoveItem A
        End If
    Next A
    
    ' ELIMINAMOS LAS FILAS EN BLANCO DEL CONTROL Fg2
    For A = 1 To Fg2.Rows - 1
        If Fg2.TextMatrix(A, 1) = "" Then
            Fg2.RemoveItem A
        End If
    Next A
    
On Error GoTo LaCague
    xCon.BeginTrans
    If QueHace = 1 Then
        xId = HallaCodigoTabla("com_ordenreq", xCon, "id")
    Else
        xId = RstLista("id")
    End If
    
    'Columna    | Descripcion
    '------------------------
    '0          | campo
    '1          | Valor
    '2          | requerido
    '3          | tipo
    '4          | msj error
    '5          | INDICA QUE EL CAMPO ES INDICE Y NO SE ESCRIBIRA CUANDO SE MODIFIQUE EL REGISTRO
    '--------------------------------
    'GRABAMOS LA CABECERA DE LA ORDEN DE REQUERIMIENTO
    xCampos(0, 0) = "id":           xCampos(0, 1) = Str(xId):                   xCampos(0, 2) = "S":    xCampos(0, 3) = "N":    xCampos(0, 4) = "":                                                                     xCampos(0, 5) = "S"
    xCampos(1, 0) = "idtipdoc":     xCampos(1, 1) = "106":                      xCampos(1, 2) = "S":    xCampos(1, 3) = "N":    xCampos(1, 4) = "":                                                                     xCampos(1, 5) = ""
    xCampos(2, 0) = "numser":       xCampos(2, 1) = NulosC(TxtNumSer.Text):     xCampos(2, 2) = "S":    xCampos(2, 3) = "C":    xCampos(2, 4) = "":                                                                     xCampos(2, 5) = ""
    xCampos(3, 0) = "numdoc":       xCampos(3, 1) = NulosC(TxtNumDoc.Text):     xCampos(3, 2) = "S":    xCampos(3, 3) = "C":    xCampos(3, 4) = "":                                                                     xCampos(3, 5) = ""
    xCampos(4, 0) = "idarea":       xCampos(4, 1) = NulosC(TxtIdArea.Text):     xCampos(4, 2) = "S":    xCampos(4, 3) = "N":    xCampos(4, 4) = "No ha especificado el area que solicita el requerimiento":             xCampos(4, 5) = ""
    xCampos(5, 0) = "idsol":        xCampos(5, 1) = TxtidSol.Text:              xCampos(5, 2) = "S":    xCampos(5, 3) = "N":    xCampos(5, 4) = "No ha especificado el solicitante del requerimiento":                  xCampos(5, 5) = ""
    xCampos(6, 0) = "fchemi":       xCampos(6, 1) = TxtFchEmi.Valor:            xCampos(6, 2) = "S":    xCampos(6, 3) = "F":    xCampos(6, 4) = "No ha especificado la fecha de emision de la orden de requerimiento":  xCampos(6, 5) = ""
    xCampos(7, 0) = "fchent":       xCampos(7, 1) = TxtFchEnt.Valor:            xCampos(7, 2) = "S":    xCampos(7, 3) = "F":    xCampos(7, 4) = "No ha especificado la fecha de entrega de la orden de requerimiento":  xCampos(7, 5) = ""
    xCampos(8, 0) = "idmon":        xCampos(8, 1) = NulosC(TxtIdMon.Text):      xCampos(8, 2) = "S":    xCampos(8, 3) = "N":    xCampos(8, 4) = "No ha especificado la moneda de la orden de requerimiento":            xCampos(8, 5) = ""
    xCampos(9, 0) = "idest":        xCampos(9, 1) = 1:                          xCampos(9, 2) = "S":    xCampos(9, 3) = "N":    xCampos(9, 4) = ""
    xCampos(10, 0) = "idsit":       xCampos(10, 1) = 1:                         xCampos(10, 2) = "S":   xCampos(10, 3) = "N":   xCampos(10, 4) = ""
    
    If QueHace = 1 Then
        If EscribirNuevoRegistro(xCampos, "com_ordenreq", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    Else
        ' ELIMINAMOS LOS DETALLES DE LA ORDEN DE REQUERIMIENTO
        xCon.Execute "DELETE * FROM com_ordenreqdet WHERE idor = " & RstLista("id") & ""
        xCon.Execute "DELETE * FROM com_ordenreqpro WHERE idor = " & RstLista("id") & ""
        xCon.Execute "DELETE * FROM com_ordenreqprocot WHERE idor = " & RstLista("id") & ""
        ' MODIFICAMOS EL REGISTRO
        If ModificarRegistro(xCampos, "com_ordenreq", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    End If
    
    ' GRABAMOS EL DETALLE DE LA ORDEN DE REQUERIMIENTO
    For A = 1 To Fg1.Rows - 1
        xCampos2(0, 0) = "idor":           xCampos2(0, 1) = Str(xId):                   xCampos2(0, 2) = "S":    xCampos2(0, 3) = "N":    xCampos2(0, 4) = "":      xCampos2(0, 5) = ""
        xCampos2(1, 0) = "iditem":         xCampos2(1, 1) = Fg1.TextMatrix(A, 7):       xCampos2(1, 2) = "S":    xCampos2(1, 3) = "N":    xCampos2(1, 4) = "":      xCampos2(1, 5) = ""
        xCampos2(2, 0) = "idcencos":       xCampos2(2, 1) = Fg1.TextMatrix(A, 9):       xCampos2(2, 2) = "N":    xCampos2(2, 3) = "N":    xCampos2(2, 4) = "":      xCampos2(2, 5) = ""
        xCampos2(3, 0) = "idunimed":       xCampos2(3, 1) = Fg1.TextMatrix(A, 8):       xCampos2(3, 2) = "S":    xCampos2(3, 3) = "N":    xCampos2(3, 4) = "":      xCampos2(3, 5) = ""
        xCampos2(4, 0) = "cantidad":       xCampos2(4, 1) = Fg1.TextMatrix(A, 4):       xCampos2(4, 2) = "S":    xCampos2(4, 3) = "N":    xCampos2(4, 4) = "":      xCampos2(4, 5) = ""
        xCampos2(5, 0) = "corr":           xCampos2(5, 1) = Fg1.TextMatrix(A, 1):       xCampos2(5, 2) = "S":    xCampos2(5, 3) = "N":    xCampos2(5, 4) = "":      xCampos2(5, 5) = ""
        
        If EscribirNuevoRegistro(xCampos2, "com_ordenreqdet", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    Next A
    
    ' GRABAMOS LOS PROVEEDORES RECOMENDADOS PARA COTIZACION
    For A = 1 To Fg2.Rows - 1
        xCampos3(0, 0) = "idor":           xCampos3(0, 1) = Str(xId):                   xCampos3(0, 2) = "S":    xCampos3(0, 3) = "N":    xCampos3(0, 4) = "":      xCampos3(0, 5) = "":
        xCampos3(1, 0) = "idpro":          xCampos3(1, 1) = Fg2.TextMatrix(A, 3):       xCampos3(1, 2) = "S":    xCampos3(1, 3) = "N":    xCampos3(1, 4) = "Proveedor incorrecto":      xCampos3(1, 5) = "":
        
        If EscribirNuevoRegistro(xCampos3, "com_ordenreqpro", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    Next A
    
    Dim xIdPro As Integer
    Dim C As Integer
    
    ' GRABAMOS EL PLAN DE COTIZACION
    For A = 1 To Fg3.Rows - 1
        
        For B = 3 To Fg3.Cols - 1
            If NulosN(Fg3.TextMatrix(A, B)) = -1 Then
                ' OBTENEMOS EL ID DEL PROVEEDOR
                For C = 1 To Fg2.Rows - 1
                    If Fg2.TextMatrix(C, 2) = Fg3.TextMatrix(0, B) Then
                        xIdPro = Fg2.TextMatrix(C, 3)
                        Exit For
                    End If
                Next C
                
                xCampos4(0, 0) = "idor":           xCampos4(0, 1) = Str(xId):                   xCampos4(0, 2) = "S":    xCampos4(0, 3) = "N":    xCampos4(0, 4) = "":
                xCampos4(1, 0) = "idpro":          xCampos4(1, 1) = xIdPro:                     xCampos4(1, 2) = "S":    xCampos4(1, 3) = "N":    xCampos4(1, 4) = ""
                xCampos4(2, 0) = "idite":          xCampos4(2, 1) = Fg3.TextMatrix(A, 2):       xCampos4(2, 2) = "S":    xCampos4(2, 3) = "N":    xCampos4(2, 4) = ""
            
                If EscribirNuevoRegistro(xCampos4, "com_ordenreqprocot", xCon) = False Then
                    xCon.RollbackTrans
                    Exit Function
                End If
            End If
        Next B
    Next A
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    
    xCon.CommitTrans
    MsgBox "El registro se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function

LaCague:
    'Resume
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el registro por el siguiente motivo: " & Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = False
End Function

Sub MuestraSegundoTab()
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    Dim B As Integer
    
    TabOne2.CurrTab = 0
    TxtNumSer.Text = NulosC(RstLista("numser"))
    TxtNumDoc.Text = NulosC(RstLista("numdoc"))
    TxtidSol.Text = RstLista("idsol")
    LblSolicitante.Caption = NulosC(RstLista("apenom"))
    TxtIdArea.Text = RstLista("idarea")
    LblArea.Caption = NulosC(RstLista("descarea"))
    TxtFchEmi.Valor = Format(RstLista("fchemi"), "dd/mm/yyyy")
    TxtFchEnt.Valor = Format(RstLista("fchent"), "dd/mm/yyyy")
    TxtIdMon.Text = RstLista("idmon")
    LblMoneda.Caption = RstLista("descmon")
    
    RST_Busq Rst, "SELECT com_ordenreqdet.*, alm_inventario.descripcion AS descitem, con_centrocosto.descripcion AS desccosto, mae_unidades.abrev AS descunimed, " _
        & " con_centrocosto.codigo FROM ((com_ordenreqdet LEFT JOIN alm_inventario ON com_ordenreqdet.iditem = alm_inventario.id) LEFT JOIN con_centrocosto ON " _
        & " com_ordenreqdet.idcencos = con_centrocosto.id) LEFT JOIN mae_unidades ON com_ordenreqdet.idunimed = mae_unidades.id WHERE (((com_ordenreqdet.idor)=" & NulosN(RstLista("id")) & "))", xCon
    
    Fg1.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("corr")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("descitem")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Rst("descunimed")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(Rst("cantidad"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(Rst("codigo"))
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(Rst("desccosto"))
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Rst("iditem")
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = Rst("idunimed")
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosN(Rst("idcencos"))
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    Else
        'Fg1.Rows = Fg1.Rows + 1
    End If
    
    Set Rst = Nothing
    
    ' MUESTRA LOS PROVEEDORES
    RST_Busq Rst, "SELECT com_ordenreqpro.idor, mae_prov.nombre, mae_prov.numruc, com_ordenreqpro.idpro FROM com_ordenreqpro LEFT JOIN mae_prov " _
        & " ON com_ordenreqpro.idpro = mae_prov.id Where (((com_ordenreqpro.idor) = " & NulosN(RstLista("id")) & ")) ORDER BY mae_prov.nombre", xCon

    Fg2.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = Rst("numruc")
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = Rst("nombre")
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = Rst("idpro")
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    Else
        'Fg2.Rows = Fg2.Rows + 1
    End If
    
    Set Rst = Nothing
    
    ' MUESTRA LOS ITEMS QUE SE COTIZARAN
    RST_Busq Rst, "SELECT com_ordenreqprocot.idpro, com_ordenreqprocot.idite, mae_prov.nombre FROM com_ordenreqprocot LEFT JOIN mae_prov " _
        & " ON com_ordenreqprocot.idpro = mae_prov.id WHERE (((com_ordenreqprocot.idor) = " & NulosN(RstLista("id")) & "))", xCon

    Fg3.Rows = 3
    CmdConfCot_Click
    
    If Rst.RecordCount <> 0 Then
        For A = 1 To Fg3.Rows - 1
            For B = 3 To Fg3.Cols - 1
                Rst.Filter = "idite = " & NulosN(Fg3.TextMatrix(A, 2)) & " AND nombre = '" & Fg3.TextMatrix(0, B) & "'"
                If Rst.RecordCount = 1 Then
                     Fg3.TextMatrix(A, B) = -1
                End If
            Next B
        Next A
    End If
    Fg1.SelectionMode = flexSelectionByRow
    Fg2.SelectionMode = flexSelectionByRow
    Fg3.SelectionMode = flexSelectionByRow
End Sub

