VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmPlanProduccion3 
   Caption         =   "Produccion - Plan de Produccion"
   ClientHeight    =   9045
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9045
   ScaleWidth      =   13575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmProgreso 
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   1065
      Left            =   4950
      TabIndex        =   4
      Top             =   8250
      Visible         =   0   'False
      Width           =   5625
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   90
         TabIndex        =   5
         Top             =   645
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
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
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5610
         X2              =   5610
         Y1              =   15
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
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   5610
         Y1              =   1050
         Y2              =   1050
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
         TabIndex        =   7
         Top             =   420
         Width           =   1575
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
         TabIndex        =   6
         Top             =   75
         Width           =   1575
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7380
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   12330
      _cx             =   21749
      _cy             =   13017
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
      Begin SizerOneLibCtl.ElasticOne Eo2 
         Height          =   6960
         Left            =   45
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   375
         Width           =   12240
         _cx             =   21590
         _cy             =   12277
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
         Appearance      =   4
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   12648447
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   3
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmPlanProduccion3.frx":0000
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   5685
            Left            =   90
            TabIndex        =   25
            Top             =   1185
            Width           =   12060
            _cx             =   21272
            _cy             =   10028
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
            ForeColor       =   -2147483633
            FrontTabColor   =   -2147483633
            BackTabColor    =   -2147483632
            TabOutlineColor =   -2147483633
            FrontTabForeColor=   -2147483630
            Caption         =   "   &Terminado   |   &Intermedios   "
            Align           =   0
            CurrTab         =   1
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
            Begin VB.Frame Frame15 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'None
               Caption         =   "Frame15"
               ForeColor       =   &H00C0C0C0&
               Height          =   5310
               Left            =   12705
               TabIndex        =   28
               Top             =   45
               Width           =   11970
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
                  TabIndex        =   30
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
                  TabIndex        =   29
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
            Begin VSFlex7Ctl.VSFlexGrid Fg2 
               Height          =   5310
               Left            =   45
               TabIndex        =   27
               Top             =   45
               Width           =   11970
               _cx             =   21114
               _cy             =   9366
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
               BackColorSel    =   64
               ForeColorSel    =   65535
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
               FormatString    =   $"FrmPlanProduccion3.frx":004E
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
            Begin VSFlex7Ctl.VSFlexGrid Fg1 
               Height          =   5310
               Left            =   -12615
               TabIndex        =   26
               Top             =   45
               Width           =   11970
               _cx             =   21114
               _cy             =   9366
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
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmPlanProduccion3.frx":0106
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
         Begin SizerOneLibCtl.ElasticOne Eo3 
            Height          =   735
            Left            =   90
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   390
            Width           =   12060
            _cx             =   21273
            _cy             =   1296
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
            Appearance      =   4
            MousePointer    =   0
            _ConvInfo       =   1
            Version         =   700
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   1
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   2
            PicturePos      =   4
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   1
            GridCols        =   3
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmPlanProduccion3.frx":01A0
            Begin VB.Frame Frame2 
               BackColor       =   &H00FFFFC0&
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   555
               Left            =   6645
               TabIndex        =   13
               Top             =   90
               Width           =   5325
               Begin VB.CommandButton CmdAdd 
                  Caption         =   "Agregar Plan de Ventas"
                  Height          =   525
                  Left            =   1275
                  TabIndex        =   22
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   1185
               End
               Begin VB.CommandButton CmdAddProd 
                  Caption         =   "Agregar Plan de Produccion"
                  Height          =   525
                  Left            =   2475
                  TabIndex        =   21
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   1185
               End
               Begin VB.CommandButton CmdVerEst 
                  Caption         =   "&Ver Estacionalidad"
                  Height          =   525
                  Left            =   75
                  TabIndex        =   20
                  Top             =   105
                  Visible         =   0   'False
                  Width           =   1185
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  Caption         =   "Nº Registros : "
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Left            =   3900
                  TabIndex        =   24
                  Top             =   165
                  Width           =   1020
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
                  Left            =   3900
                  TabIndex        =   23
                  Top             =   375
                  Width           =   1020
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   555
               Left            =   90
               TabIndex        =   12
               Top             =   90
               Width           =   6360
               Begin VB.TextBox TxtDesc 
                  Appearance      =   0  'Flat
                  Height          =   300
                  Left            =   945
                  Locked          =   -1  'True
                  TabIndex        =   14
                  Text            =   "TxtDesc"
                  Top             =   75
                  Width           =   5385
               End
               Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
                  Height          =   300
                  Left            =   945
                  TabIndex        =   15
                  Top             =   390
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
                  Left            =   4650
                  TabIndex        =   16
                  Top             =   390
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
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "Fch. Termino"
                  Height          =   195
                  Left            =   3465
                  TabIndex        =   19
                  Top             =   420
                  Width           =   930
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Descripcion"
                  Height          =   195
                  Left            =   45
                  TabIndex        =   18
                  Top             =   105
                  Width           =   840
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  Caption         =   "Fch. Inicio"
                  Height          =   195
                  Left            =   45
                  TabIndex        =   17
                  Top             =   420
                  Width           =   735
               End
            End
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Plan de Produccion"
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
            Height          =   240
            Left            =   90
            TabIndex        =   10
            Top             =   90
            Width           =   12060
         End
      End
      Begin SizerOneLibCtl.ElasticOne Eo1 
         Height          =   6960
         Left            =   -12885
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   375
         Width           =   12240
         _cx             =   21590
         _cy             =   12277
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
         Appearance      =   4
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   12640511
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   2
         GridCols        =   1
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmPlanProduccion3.frx":01ED
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6480
            Left            =   90
            TabIndex        =   9
            Top             =   390
            Width           =   12060
            _ExtentX        =   21273
            _ExtentY        =   11430
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nº Proyecto"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descripcion"
            Columns(1).DataField=   "descripcion"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Ini"
            Columns(2).DataField=   "fchini"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch. Fin"
            Columns(3).DataField=   "fchfin"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Estado"
            Columns(4).DataField=   "estado"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2037"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1958"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=8202"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=8123"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1826"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1746"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1799"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1720"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=1667"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1588"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HDBFDFD&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H400000&"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Named:id=33:Normal"
            _StyleDefs(57)  =   ":id=33,.parent=0"
            _StyleDefs(58)  =   "Named:id=34:Heading"
            _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(60)  =   ":id=34,.wraptext=-1"
            _StyleDefs(61)  =   "Named:id=35:Footing"
            _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(63)  =   "Named:id=36:Selected"
            _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(65)  =   "Named:id=37:Caption"
            _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(67)  =   "Named:id=38:HighlightRow"
            _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=39:EvenRow"
            _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(71)  =   "Named:id=40:OddRow"
            _StyleDefs(72)  =   ":id=40,.parent=33"
            _StyleDefs(73)  =   "Named:id=41:RecordSelector"
            _StyleDefs(74)  =   ":id=41,.parent=34"
            _StyleDefs(75)  =   "Named:id=42:FilterBar"
            _StyleDefs(76)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Plan de Produccion"
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
            Height          =   240
            Left            =   90
            TabIndex        =   3
            Top             =   90
            Width           =   12060
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12660
      Top             =   90
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
            Picture         =   "FrmPlanProduccion3.frx":0230
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanProduccion3.frx":0774
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanProduccion3.frx":0B06
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanProduccion3.frx":0C8A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanProduccion3.frx":10DE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanProduccion3.frx":11F6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanProduccion3.frx":173A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanProduccion3.frx":1C7E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanProduccion3.frx":1D92
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanProduccion3.frx":1EA6
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanProduccion3.frx":22FA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanProduccion3.frx":2466
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanProduccion3.frx":29AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanProduccion3.frx":2CC8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   13575
      _ExtentX        =   23945
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
   End
End
Attribute VB_Name = "FrmPlanProduccion3"
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

Private Sub CmdAdd_Click()
    ' EJECUTA LA BUSQUEDA DE UN PLAN DE VENTAS
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCodItem As String
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":       xCampos(1, 1) = "id":            xCampos(1, 2) = "2000":    xCampos(1, 3) = "N"

    xform.SQLCad = "SELECT ges_planventas.id, ges_planventas.descripcion , ges_planventas.fchini, ges_planventas.fchfin From ges_planventas " _
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
        If IsDate(xRs("fchini")) = True Then
            MostrarPlanVentas xId, Month(xRs("fchini")), Year(xRs("fchini"))
            MostrarIntermedios xId, 4, 2011
            LblNumReg.Caption = Fg1.Rows - 1
            DoEvents
        End If
        Set xRs = Nothing
        PintarGrid
        TabOne2.CurrTab = 0
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Sub PintarGrid()
    Dim A As Integer
    GRID_COLOR_FONDO Fg1, 1, Fg1.Cols - 5, Fg1.Rows - 1, Fg1.Cols - 5, &HFFFFC0, flexFillRepeat
    GRID_COLOR_FONDO Fg1, 1, Fg1.Cols - 4, Fg1.Rows - 1, Fg1.Cols - 2, &HC0FFC0, flexFillRepeat

    GRID_COLOR_FONDO Fg2, 1, Fg2.Cols - 5, Fg2.Rows - 1, Fg2.Cols - 5, &HFFFFC0, flexFillRepeat
    GRID_COLOR_FONDO Fg2, 1, Fg2.Cols - 4, Fg2.Rows - 1, Fg2.Cols - 2, &HC0FFC0, flexFillRepeat
    
    ' ALINEAMOS LOS ENCABEZADOS DELAS COLUMNAS
    For A = 1 To Fg1.Cols - 1
        Fg1.FixedAlignment(A) = flexAlignCenterCenter
        FORMATO_CELDA Fg1, 0, A, , True, &H8000000F, Fg1.TextMatrix(0, A)
    Next A
    
    For A = 1 To Fg2.Cols - 1
        Fg2.FixedAlignment(A) = flexAlignCenterCenter
        FORMATO_CELDA Fg2, 0, A, , True, &H8000000F, Fg2.TextMatrix(0, A)
    Next A
   
End Sub

Sub MostrarIntermedios(IdPlanVentas As Integer, xMesInicio As Integer, xAñoInicio As Integer)
    Dim xSQL As String
    
    xSQL = "SELECT todo.iditem, todo.descripcion, todo.idunimed, todo.abrev, Sum(todo.total) AS SumaDetotal, materiaprima.id_matprima, materiaprima.des_matpri " _
        & " FROM " _
        & " ( " _
        & "     SELECT pro_recetains.iditem, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, Sum([productos]![totpro]*[pro_recetains]![canpro]) AS total " _
        & "     FROM (( " _
        & "     ( " _
        & "         SELECT ges_planventasdet.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, Sum(ges_planventasdet.cantidad) AS totpro " _
        & "         FROM (ges_planventasdet LEFT JOIN alm_inventario ON ges_planventasdet.codpro = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        & "         GROUP BY ges_planventasdet.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, ges_planventasdet.idpv " _
        & "         Having (((ges_planventasdet.idpv) = " & IdPlanVentas & ")) ORDER BY alm_inventario.descripcion " _
        & "     ) AS productos " _
        & "     LEFT JOIN pro_receta ON productos.codpro = pro_receta.iditem) LEFT JOIN (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) " _
        & "     ON pro_receta.id = pro_recetains.idrec) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
        & "     GROUP BY pro_recetains.iditem, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, alm_inventario.tippro, pro_receta.prirec " _
        & "     Having (((alm_inventario.tippro) = 3) And ((pro_receta.prirec) = 1)) "
    
    xSQL = xSQL _
        & "     UNION " _
        & "     SELECT pro_recetains.iditem, alm_inventario.descripcion, pro_recetains.idunimed, mae_unidades.abrev, Sum([canpro]*pro_nivel1.total) AS total " _
        & "     FROM ((( " _
        & "     ( " _
        & "         SELECT pro_recetains.iditem, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, Sum([productos]![totpro]*[pro_recetains]![canpro]) AS total " _
        & "         FROM (( " _
        & "         ( " _
        & "             SELECT ges_planventasdet.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, Sum(ges_planventasdet.cantidad) AS totpro " _
        & "             FROM (ges_planventasdet LEFT JOIN alm_inventario ON ges_planventasdet.codpro = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        & "             GROUP BY ges_planventasdet.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, ges_planventasdet.idpv " _
        & "             Having (((ges_planventasdet.idpv) = " & IdPlanVentas & ")) ORDER BY alm_inventario.descripcion " _
        & "         ) AS productos " _
        & "         LEFT JOIN pro_receta ON productos.codpro = pro_receta.iditem) LEFT JOIN (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) " _
        & "         ON pro_receta.id = pro_recetains.idrec) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
        & "         GROUP BY pro_recetains.iditem, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, alm_inventario.tippro, pro_receta.prirec " _
        & "         Having (((alm_inventario.tippro) = 3) And ((pro_receta.prirec) = 1))" _
        & "     ) AS pro_nivel1 LEFT JOIN pro_receta ON pro_nivel1.iditem = pro_receta.iditem) LEFT JOIN pro_recetains ON pro_receta.id = pro_recetains.idrec) " _
        & "     LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
        & "     GROUP BY pro_recetains.iditem, alm_inventario.descripcion, pro_recetains.idunimed, mae_unidades.abrev, alm_inventario.tippro, pro_receta.prirec " _
        & "     Having (((alm_inventario.tippro) = 3) And ((pro_receta.prirec) = 1)) "
    
    xSQL = xSQL _
        & "     UNION " _
        & "     SELECT pro_recetains.iditem, alm_inventario.descripcion, pro_recetains.idunimed, mae_unidades.abrev, Sum([canpro]*segundonivel.total) AS total " _
        & "     FROM ( " _
        & "     ( " _
        & "         SELECT pro_recetains.iditem, alm_inventario.descripcion, pro_recetains.idunimed, mae_unidades.abrev, Sum([canpro]*pro_nivel1.total) AS total " _
        & "         FROM ((( " _
        & "         ( " _
        & "             SELECT pro_recetains.iditem, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, Sum([productos]![totpro]*[pro_recetains]![canpro]) AS total " _
        & "             FROM (( " _
        & "             ( " _
        & "                 SELECT ges_planventasdet.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, Sum(ges_planventasdet.cantidad) AS totpro " _
        & "                 FROM (ges_planventasdet LEFT JOIN alm_inventario ON ges_planventasdet.codpro = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        & "                 GROUP BY ges_planventasdet.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, ges_planventasdet.idpv " _
        & "                 Having (((ges_planventasdet.idpv) = " & IdPlanVentas & ")) ORDER BY alm_inventario.descripcion " _
        & "             ) AS productos " _
        & "             LEFT JOIN pro_receta ON productos.codpro = pro_receta.iditem) LEFT JOIN (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) " _
        & "             ON pro_receta.id = pro_recetains.idrec) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
        & "             GROUP BY pro_recetains.iditem, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, alm_inventario.tippro, pro_receta.prirec " _
        & "             Having (((alm_inventario.tippro) = 3) And ((pro_receta.prirec) = 1)) " _
        & "         ) AS pro_nivel1 " _
        & "         LEFT JOIN pro_receta ON pro_nivel1.iditem = pro_receta.iditem) LEFT JOIN pro_recetains ON pro_receta.id = pro_recetains.idrec) LEFT JOIN alm_inventario " _
        & "         ON pro_recetains.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
        & "         GROUP BY pro_recetains.iditem, alm_inventario.descripcion, pro_recetains.idunimed, mae_unidades.abrev, alm_inventario.tippro, pro_receta.prirec " _
        & "         Having (((alm_inventario.tippro) = 3) And ((pro_receta.prirec) = 1)) "
        
    xSQL = xSQL _
        & "     ) AS segundonivel " _
        & "     LEFT JOIN (pro_receta LEFT JOIN (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) ON pro_receta.id = pro_recetains.idrec) " _
        & "     ON segundonivel.iditem = pro_receta.iditem) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
        & "     GROUP BY pro_recetains.iditem, alm_inventario.descripcion, pro_recetains.idunimed, mae_unidades.abrev, alm_inventario.tippro, pro_receta.prirec " _
        & "     Having (((alm_inventario.tippro) = 3) And ((pro_receta.prirec) = 1)) " _
        & " ) AS todo " _
        & " LEFT JOIN " _
        & " ( " _
        & "     SELECT pro_receta.iditem AS codpro, alm_inventario_1.descripcion AS pro_descripcion, pro_recetains.iditem AS id_matprima, alm_inventario.descripcion AS des_matpri " _
        & "     FROM (((pro_receta LEFT JOIN pro_recetains ON pro_receta.id = pro_recetains.idrec) LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) " _
        & "     LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) LEFT JOIN alm_inventario AS alm_inventario_1 ON pro_receta.iditem = alm_inventario_1.id " _
        & "     Where (((alm_inventario.tippro) = 1) And ((pro_receta.prirec) = 1)) " _
        & " ) AS materiaprima " _
        & " ON todo.iditem = materiaprima.codpro GROUP BY todo.iditem, todo.descripcion, todo.idunimed, todo.abrev, materiaprima.id_matprima, materiaprima.des_matpri " _
        & " ORDER BY todo.descripcion"
       
    
    
    Dim xRstInter As New ADODB.Recordset
    Dim A, xColIni, B, xNumMesEst As Integer
    Dim xTotMes As Double
    
    RST_Busq xRstInter, xSQL, xCon
    
    Fg2.Rows = 1
    Fg2.Cols = 6
    
    If xRstInter.RecordCount <> 0 Then
        
        xRstInter.MoveFirst
        For A = 1 To xRstInter.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = xRstInter("iditem")
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = xRstInter("idunimed")
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = NulosN(xRstInter("id_matprima"))
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = xRstInter("descripcion")
            Fg2.TextMatrix(Fg2.Rows - 1, 5) = xRstInter("abrev")
            xRstInter.MoveNext
            If xRstInter.EOF = True Then Exit For
        Next A
        
        xColIni = 6

        For A = xMesInicio To 12
            Fg2.Cols = Fg2.Cols + 1
            Fg2.TextMatrix(0, Fg2.Cols - 1) = Format(A, "00") & "-" & Format(xAñoInicio, "0000")
        Next A

        For A = 1 To xMesInicio - 1
            Fg2.Cols = Fg2.Cols + 1
            Fg2.TextMatrix(0, Fg2.Cols - 1) = Format(A, "00") & "-" & Format(xAñoInicio + 1, "0000")
        Next A
        
        Fg2.Cols = Fg2.Cols + 1
        Fg2.TextMatrix(0, Fg2.Cols - 1) = "Programado"
        Fg2.ColWidth(Fg2.Cols - 1) = 1200
        
        ' AGREGAMOS LOS TOTALES Y LAS CANTIDADES POR MESES LOS PRODUCTOS QUE NO TIENEN MATERIA PRIMA
        xRstInter.MoveFirst
        For A = 1 To xRstInter.RecordCount
            Fg2.TextMatrix(A, Fg2.Cols - 1) = Format(NulosN(xRstInter("SumaDetotal")), "0.00")
            
            xColIni = 6
            xTotMes = 0
            If NulosN(xRstInter("id_matprima")) = 0 Then
                xTotMes = (NulosN(xRstInter("SumaDetotal") / 12))
                For B = 6 To Fg2.Cols - 2
                    Fg2.TextMatrix(A, B) = Format(NulosN(xTotMes), "0.00")
                    xColIni = xColIni + 1
                Next B
            Else
                xNumMesEst = HallarNumMesEstacionalidad(NulosN(xRstInter("id_matprima")))
                If xNumMesEst <> 0 Then
                    xTotMes = (NulosN(xRstInter("SumaDetotal") / xNumMesEst))
                    
                    For B = 6 To Fg2.Cols - 2
                        'Fg2.TextMatrix(A, B) = Format(NulosN(xTotMes), "0.00")
                        If AplicaEstacionalidad(NulosN(xRstInter("id_matprima")), NulosN(Mid(Fg2.TextMatrix(0, B), 1, 2))) = True Then
                            Fg2.TextMatrix(A, B) = Format(NulosN(xTotMes), "0.00")
                        End If
                        xColIni = xColIni + 1
                    Next B
                End If
            End If
            
            xRstInter.MoveNext
            
            If xRstInter.EOF = True Then Exit For
        Next A
        
        
        ' ESCRIBIMOS LOS TOTALES
        Dim xStkIni, xTotPro, xTotal As Double
        Dim AnoTra As Integer
        Dim RstTodProd As New Recordset
        
        Fg2.Cols = Fg2.Cols + 4
        
        Fg2.TextMatrix(0, Fg2.Cols - 4) = "Stock Ini"
        Fg2.TextMatrix(0, Fg2.Cols - 3) = "Producido"
        Fg2.TextMatrix(0, Fg2.Cols - 2) = "Total"
        Fg2.TextMatrix(0, Fg2.Cols - 1) = "Diferencia"
        Fg2.ColWidth(Fg2.Cols - 1) = 1100
        
        AnoTra = Year(Now)
                
        For A = 1 To Fg2.Rows - 1
            xStkIni = SaldoActual(NulosN(Fg2.TextMatrix(A, 1)), "01/01/" & Format(AnoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)
            xTotPro = HallarTotalProducido(NulosN(Fg2.TextMatrix(A, 1)), TxtFchIni.Valor)
            
            Fg2.TextMatrix(A, Fg2.Cols - 4) = Format(xStkIni, "0.00")
            Fg2.TextMatrix(A, Fg2.Cols - 3) = Format(xTotPro, "0.00")
            Fg2.TextMatrix(A, Fg2.Cols - 2) = Format(xTotPro + xStkIni, "0.00")
            
            xTotal = ((xTotPro + xStkIni) - NulosN(Fg2.TextMatrix(A, Fg2.Cols - 5)))
            'Fg2.TextMatrix(A, Fg2.Cols - 1) = Format(xTotal, "0.00")
            If xTotal > 0 Then
                FORMATO_CELDA Fg2, CLng(A), Fg2.Cols - 1, &HFF0000, True, , Format(xTotal, "0.00")
            Else
                FORMATO_CELDA Fg2, CLng(A), Fg2.Cols - 1, &HC0&, True, , Format(xTotal, "0.00")
            End If
        Next A
        
        Fg2.FrozenCols = 5
    End If
End Sub

Function HallarTotalProducido(xIdProducto As Integer, Desde As String) As Double
    Dim xRst As New ADODB.Recordset
    Dim xSQL As String
    
    xSQL = "SELECT Sum(pro_producciondet.cantidad) AS SumaDecantidad FROM pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro " _
        & " WHERE (((pro_produccion.dia)>=CDate('" & Desde & "'))) GROUP BY pro_producciondet.iditem HAVING (((pro_producciondet.iditem)=" & xIdProducto & "))"

    RST_Busq xRst, xSQL, xCon
    
    If xRst.RecordCount <> 0 Then
        HallarTotalProducido = xRst("SumaDecantidad")
    Else
        HallarTotalProducido = 0
    End If
End Function

Function AplicaEstacionalidad(IdMateriaPrima As Integer, IdMes As Integer) As Boolean
    Dim xRst As New ADODB.Recordset
    Dim xSQL As String
    
    xSQL = "SELECT pro_estacionalidad.iditem, pro_estacionalidad.ene, pro_estacionalidad.feb, pro_estacionalidad.mar, pro_estacionalidad.abr, pro_estacionalidad.may, " _
        & " pro_estacionalidad.jun, pro_estacionalidad.jul, pro_estacionalidad.ago, pro_estacionalidad.set, pro_estacionalidad.oct, pro_estacionalidad.nov, " _
        & " pro_estacionalidad.dic " _
        & " From pro_estacionalidad WHERE (((pro_estacionalidad.iditem)=" & IdMateriaPrima & "))"

    RST_Busq xRst, xSQL, xCon
        
    AplicaEstacionalidad = False
    
    If xRst.RecordCount <> 0 Then
        If IdMes = 1 Then
            If xRst("ene") = 2 Or xRst("ene") = 3 Then AplicaEstacionalidad = True
        End If
        
        If IdMes = 2 Then
            If xRst("feb") = 2 Or xRst("feb") = 3 Then AplicaEstacionalidad = True
        End If
        
        If IdMes = 3 Then
            If xRst("mar") = 2 Or xRst("mar") = 3 Then AplicaEstacionalidad = True
        End If
        
        If IdMes = 4 Then
            If xRst("abr") = 2 Or xRst("abr") = 3 Then AplicaEstacionalidad = True
        End If
        
        If IdMes = 5 Then
            If xRst("may") = 2 Or xRst("may") = 3 Then AplicaEstacionalidad = True
        End If
        
        If IdMes = 6 Then
            If xRst("jun") = 2 Or xRst("jun") = 3 Then AplicaEstacionalidad = True
        End If
        
        If IdMes = 7 Then
            If xRst("jul") = 2 Or xRst("jul") = 3 Then AplicaEstacionalidad = True
        End If
        
        If IdMes = 8 Then
            If xRst("ago") = 2 Or xRst("ago") = 3 Then AplicaEstacionalidad = True
        End If
        
        If IdMes = 9 Then
            If xRst("set") = 2 Or xRst("set") = 3 Then AplicaEstacionalidad = True
        End If
        
        If IdMes = 10 Then
            If xRst("oct") = 2 Or xRst("oct") = 3 Then AplicaEstacionalidad = True
        End If
        
        If IdMes = 11 Then
            If xRst("nov") = 2 Or xRst("nov") = 3 Then AplicaEstacionalidad = True
        End If
        
        If IdMes = 12 Then
            If xRst("dic") = 2 Or xRst("dic") = 3 Then AplicaEstacionalidad = True
        End If
                
    Else
        AplicaEstacionalidad = False
    End If
End Function

Function HallarNumMesEstacionalidad(IdMateriaPrima As Integer) As Integer
    Dim xRst As New ADODB.Recordset
    Dim xSQL As String
    Dim xNumMes As Integer
    
    xSQL = "SELECT pro_estacionalidad.iditem, pro_estacionalidad.ene, pro_estacionalidad.feb, pro_estacionalidad.mar, pro_estacionalidad.abr, pro_estacionalidad.may, " _
        & " pro_estacionalidad.jun, pro_estacionalidad.jul, pro_estacionalidad.ago, pro_estacionalidad.set, pro_estacionalidad.oct, pro_estacionalidad.nov, " _
        & " pro_estacionalidad.dic " _
        & " From pro_estacionalidad WHERE (((pro_estacionalidad.iditem)=" & IdMateriaPrima & "))"

    RST_Busq xRst, xSQL, xCon
    
    xNumMes = 0
    If xRst.RecordCount <> 0 Then
        If xRst("ene") = 2 Or xRst("ene") = 3 Then xNumMes = xNumMes + 1
        If xRst("feb") = 2 Or xRst("feb") = 3 Then xNumMes = xNumMes + 1
        If xRst("mar") = 2 Or xRst("mar") = 3 Then xNumMes = xNumMes + 1
        If xRst("abr") = 2 Or xRst("abr") = 3 Then xNumMes = xNumMes + 1
        If xRst("may") = 2 Or xRst("may") = 3 Then xNumMes = xNumMes + 1
        If xRst("jun") = 2 Or xRst("jun") = 3 Then xNumMes = xNumMes + 1
        If xRst("jul") = 2 Or xRst("jul") = 3 Then xNumMes = xNumMes + 1
        If xRst("ago") = 2 Or xRst("ago") = 3 Then xNumMes = xNumMes + 1
        If xRst("set") = 2 Or xRst("set") = 3 Then xNumMes = xNumMes + 1
        If xRst("oct") = 2 Or xRst("oct") = 3 Then xNumMes = xNumMes + 1
        If xRst("nov") = 2 Or xRst("nov") = 3 Then xNumMes = xNumMes + 1
        If xRst("dic") = 2 Or xRst("dic") = 3 Then xNumMes = xNumMes + 1
        
        HallarNumMesEstacionalidad = xNumMes
    Else
        HallarNumMesEstacionalidad = 0
    End If

End Function

Sub MostrarPlanVentas(IdPlanVentas As Integer, xMesInicio As Integer, xAñoInicio As Integer)
    Dim xRstPlan As New ADODB.Recordset
    Dim xSQL As String
    Dim A, B, xColIni As Integer
    
    xSQL = "SELECT * FROM ges_planventas WHERE id = " & IdPlanVentas & ""
    RST_Busq xRstPlan, xSQL, xCon
    
    If xRstPlan.RecordCount <> 0 Then
        TxtDesc.Text = xRstPlan("descripcion")
        TxtFchIni.Valor = Format(xRstPlan("fchini"), "dd/mm/yyyy")
        TxtFchFin.Valor = Format(xRstPlan("fchfin"), "dd/mm/yyyy")
    Else
        MsgBox "EL plan de ventas especificado no existe, especifique otro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set xRstPlan = Nothing
        Exit Sub
    End If
    

    ' CREAMOS LA CONSULA PIVOT DEL PLAN DE VENTAS SELECCIONADO
    xSQL = "TRANSFORM Sum(ges_planventasdet.cantidad) AS SumaDecantidad" _
        & " SELECT ges_planventasdet.idpv, ges_planventasdet.codpro, alm_inventario.descripcion, mae_unidades.id, mae_unidades.abrev, " _
        & " Sum(ges_planventasdet.cantidad) AS total" _
        & " FROM (ges_planventasdet LEFT JOIN alm_inventario ON ges_planventasdet.codpro = alm_inventario.id) LEFT JOIN mae_unidades " _
        & " ON alm_inventario.idunimed = mae_unidades.id Where (((ges_planventasdet.idpv) = " & IdPlanVentas & ")) " _
        & " GROUP BY ges_planventasdet.idpv, ges_planventasdet.codpro, alm_inventario.descripcion, mae_unidades.id, mae_unidades.abrev " _
        & " ORDER BY alm_inventario.descripcion Pivot ges_planventasdet.idmes"

    RST_Busq xRstPlan, xSQL, xCon
    
    Fg1.Cols = 5
    Fg1.Rows = 1
    
    If xRstPlan.RecordCount <> 0 Then
        
        xRstPlan.MoveFirst
        For A = 1 To xRstPlan.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = xRstPlan("codpro")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = xRstPlan("id")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = xRstPlan("descripcion")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = xRstPlan("abrev")
            xRstPlan.MoveNext
            If xRstPlan.EOF = True Then Exit For
        Next A
        
        xColIni = 5
        
        For A = xMesInicio To 12
            Fg1.Cols = Fg1.Cols + 1
            Fg1.TextMatrix(0, Fg1.Cols - 1) = Format(A, "00") & "-" & Format(xAñoInicio, "0000")
        Next A
        
        For A = xMesInicio To 12
            xRstPlan.MoveFirst
            
            For B = 1 To Fg1.Rows - 1
                Fg1.TextMatrix(B, xColIni) = Format(xRstPlan(NulosC(A)), "0.00")
                
                xRstPlan.MoveNext
                If xRstPlan.EOF = True Then Exit For
            Next B
            xColIni = xColIni + 1
        Next A
        
                
        For A = 1 To xMesInicio - 1
            Fg1.Cols = Fg1.Cols + 1
            Fg1.TextMatrix(0, Fg1.Cols - 1) = Format(A, "00") & "-" & Format(xAñoInicio + 1, "0000")
        Next A
        
        For A = 1 To xMesInicio - 1
            xRstPlan.MoveFirst
            
            For B = 1 To Fg1.Rows - 1
                Fg1.TextMatrix(B, xColIni) = Format(xRstPlan(NulosC(A)), "0.00")
                
                xRstPlan.MoveNext
                If xRstPlan.EOF = True Then Exit For
            Next B
            xColIni = xColIni + 1
        Next A
        
        
        ' ESCRIBIMOS LOS TOTALES
        Fg1.Cols = Fg1.Cols + 5
        Fg1.TextMatrix(0, Fg1.Cols - 5) = "Programado"
        Fg1.ColWidth(Fg1.Cols - 1) = 1200
        Fg1.TextMatrix(0, Fg1.Cols - 4) = "Stock Ini"
        Fg1.TextMatrix(0, Fg1.Cols - 3) = "Producido"
        Fg1.TextMatrix(0, Fg1.Cols - 2) = "Total"
        Fg1.TextMatrix(0, Fg1.Cols - 1) = "Diferencia"
        Fg1.ColWidth(Fg1.Cols - 1) = 1200
                
        Dim xStkIni, xTotal As Double
        Dim AnoTra As Integer
        AnoTra = Year(Now)
        Dim RstTodProd As New Recordset
        
        xSQL = "SELECT ges_planventasdet.idpv, ges_planventasdet.codpro, alm_inventario.descripcion, (SELECT Sum(pro_producciondet.cantidad) AS SumaDecantidad " _
            & " FROM pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro WHERE (((pro_produccion.dia)>=CDate('" & TxtFchIni.Valor & "'))) " _
            & " GROUP BY pro_producciondet.iditem HAVING (((pro_producciondet.iditem)=ges_planventasdet.codpro))) AS totpro " _
            & " FROM ges_planventasdet LEFT JOIN alm_inventario ON ges_planventasdet.codpro = alm_inventario.id GROUP BY ges_planventasdet.idpv, ges_planventasdet.codpro, " _
            & " alm_inventario.descripcion Having (((ges_planventasdet.idpv) = " & IdPlanVentas & ")) ORDER BY alm_inventario.descripcion"

        RST_Busq RstTodProd, xSQL, xCon

        xRstPlan.MoveFirst
        RstTodProd.MoveFirst
        For A = 1 To Fg1.Rows
            xStkIni = SaldoActual(NulosN(Fg1.TextMatrix(A, 1)), "01/01/" & Format(AnoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)
            
            Fg1.TextMatrix(A, Fg1.Cols - 5) = Format(xRstPlan("total"), "0.00")
            Fg1.TextMatrix(A, Fg1.Cols - 4) = Format(xStkIni, "0.00")
            Fg1.TextMatrix(A, Fg1.Cols - 3) = Format(RstTodProd("totpro"), "0.00")
            Fg1.TextMatrix(A, Fg1.Cols - 2) = Format(RstTodProd("totpro") + xStkIni, "0.00")
            
            xTotal = ((NulosN(RstTodProd("totpro")) + xStkIni) - xRstPlan("total"))
            
            If xTotal > 0 Then
                FORMATO_CELDA Fg1, CLng(A), Fg1.Cols - 1, &HFF0000, True, , Format(xTotal, "0.00")
            Else
                FORMATO_CELDA Fg1, CLng(A), Fg1.Cols - 1, &HC0&, True, , Format(xTotal, "0.00")
            End If
            
            xRstPlan.MoveNext
            RstTodProd.MoveNext
            
            If xRstPlan.EOF = True Then Exit For
        Next A
    End If
    
    Fg1.FrozenCols = 4
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

Private Sub Dg1_DblClick()
    MuestraSegundoTab
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

Sub Bloquea()
    TxtDesc.Locked = Not TxtDesc.Locked
    TxtFchIni.Locked = Not TxtFchIni.Locked
    TxtFchFin.Locked = Not TxtFchFin.Locked
End Sub

Sub Blanquea()
    TxtDesc.Text = ""
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
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
    
    Toolbar1.Buttons(12).Enabled = Not Toolbar1.Buttons(12).Enabled
    Toolbar1.Buttons(14).Enabled = Not Toolbar1.Buttons(14).Enabled
    Toolbar1.Buttons(15).Enabled = Not Toolbar1.Buttons(15).Enabled
End Sub

Sub SetearForm()
    ' POSICIONAMOA EL FORMULARIO
    Me.Caption = "Gestion - Plan de Produccion"
    Me.Width = 12000
    Me.Height = 8200
    
    ' posicionamos el tab
    TabOne1.Left = 0
    TabOne1.Top = 360
    TabOne1.Width = Me.Width - 150
    TabOne1.Height = Me.Height - 900
    
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    
    Eo1.BackColor = &H8000000F
    Eo2.BackColor = &H8000000F
    Eo3.BackColor = &H8000000F
    
    Eo1.BorderWidth = 1
    Eo2.BorderWidth = 1
    Eo3.BorderWidth = 1
        
    Eo1.ChildSpacing = 1
    Eo2.ChildSpacing = 1
    Eo3.ChildSpacing = 1
        
    Fg1.BackColor = &HDBFDFD
    Fg2.BackColor = &HDBFDFD
    
    Fg1.ColWidth(1) = 0
    Fg1.ColWidth(2) = 0
    
    Fg2.ColWidth(1) = 0
    Fg2.ColWidth(2) = 0
    Fg2.ColWidth(3) = 0
        
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.BackColorSel = &H40&
    Fg1.ForeColorSel = &HFFFF&
    Fg1.AutoSearch = flexSearchFromTop
    Fg1.ExplorerBar = flexExSortShowAndMove
    
    Fg2.SelectionMode = flexSelectionByRow
    Fg2.BackColorSel = &H40&
    Fg2.ForeColorSel = &HFFFF&
    Fg2.AutoSearch = flexSearchFromTop
    Fg2.ExplorerBar = flexExSortShowAndMove
        
    Label1.Width = Eo1.Width - 90
End Sub
Private Sub Form_Load()

    SetearForm
    ' PRIMER EVENTO DEL FORMULARIO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    'Frame1.BackColor = &H8000000F
    'Frame2.BackColor = &H8000000F
    'Frame3.BackColor = &H8000000F
    'Frame4.BackColor = &H8000000F
    'Frame15.BackColor = &H8000000F
    
    'Fg1.AllowUserResizing = flexResizeColumns
    'Fg1.AutoSearch = flexSearchFromTop
    'Fg1.ExplorerBar = flexExSortShowAndMove
    
'    Fg2.AllowUserResizing = flexResizeColumns
'    Fg2.AutoSearch = flexSearchFromTop
'    Fg2.ExplorerBar = flexExSortShowAndMove
    
    'Fg1.ColWidth(0) = 0
    'Fg1.ColWidth(2) = 0
    
    'Fg2.ColWidth(0) = 0
    'Fg2.ColWidth(2) = 0
    
    QueHace = 3
    SeEjecuto = False
    TabOne1.CurrTab = 0
   ' TabOne2.CurrTab = 0
   ' Fg1.SelectionMode = flexSelectionByRow
   ' Fg2.SelectionMode = flexSelectionByRow
    
    'Fg1.FrozenCols = 3
    'Fg2.FrozenCols = 3
End Sub

Private Sub Form_Resize()
    CambiarTamaño
End Sub

Sub CambiarTamaño()
    If Me.WindowState = 1 Then Exit Sub
    
    TabOne1.Width = Me.Width - 150
    TabOne1.Height = Me.Height - 900
    
    Dg1.Width = Eo1.Width - 60
    Dg1.Height = Eo1.Height - 500
    
    Label1.Width = Eo1.Width - 200
End Sub

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
    Label2.Caption = "Agregando Plan de Produccion"
   
    Fg1.Cols = 5
    TxtDesc.SetFocus
End Sub

Sub Modificar()
    QueHace = 2
    xHorIni = Time
    Label1.Caption = "Modificando Plan de Produccion"
    'Blanquea
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    'PreparaRST TxtFchIni.Valor, TxtFchFin.Valor
    
    Fg1.Editable = flexEDKbdMouse
    TxtDesc.SetFocus
End Sub

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
    Dim A, B, xMes As Integer
 
    xId = 0
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
    RstCab("mesini") = NulosN(Mid(Fg1.TextMatrix(0, 5), 1, 2))
    RstCab("año") = Year(CDate(TxtFchIni.Valor))
    RstCab.Update
    
    ' GRABAMOS LOS PRODUCTOS A VENDER
    For A = 1 To Fg1.Rows - 1
        For B = 5 To Fg1.Cols - 6
            xMes = NulosN(Mid(Fg1.TextMatrix(0, B), 1, 2))
            RstDet.AddNew
            RstDet("idpv") = xId
            RstDet("codpro") = Trim(Fg1.TextMatrix(A, 1))
            RstDet("idmes") = xMes
            RstDet("cantidad") = NulosN(Fg1.TextMatrix(A, B))
            RstDet.Update
        Next B
    Next A
    
    ' GRABAMOS LOS PRODUCTOS INTERMEDIOS
    For A = 1 To Fg2.Rows - 1
        For B = 6 To Fg2.Cols - 6
            xMes = NulosN(Mid(Fg2.TextMatrix(0, B), 1, 2))
            RstDet2.AddNew
            RstDet2("idpv") = xId
            RstDet2("codpro") = Trim(Fg2.TextMatrix(A, 1))
            RstDet2("idmes") = xMes
            RstDet2("cantidad") = NulosN(Fg2.TextMatrix(A, B))
            RstDet2.Update
        Next B
    Next A
    
    
'    Dim xFila, xCol As Integer
'
'    For xFila = 1 To Fg1.Rows - 1
'        xMes = idMesIni
'        For xCol = 4 To Fg1.Cols - 5
'            RstDet.AddNew
'            RstDet("idpv") = xId
'            RstDet("codpro") = Trim(Fg1.TextMatrix(xFila, 0))
'            RstDet("idmes") = xMes
'            RstDet("cantidad") = NulosN(Fg1.TextMatrix(xFila, xCol))
'            RstDet.Update
'            xMes = xMes + 1
'            If xMes > 12 Then xMes = 1
'        Next xCol
'    Next xFila
'
'    For xFila = 1 To Fg2.Rows - 1
'        xMes = idMesIni
'        For xCol = 4 To Fg2.Cols - 5
'            RstDet2.AddNew
'            RstDet2("idpv") = xId
'            RstDet2("codpro") = Trim(Fg2.TextMatrix(xFila, 0))
'            RstDet2("idmes") = xMes
'            RstDet2("cantidad") = NulosN(Fg2.TextMatrix(xFila, xCol))
'
'            RstDet2.Update
'            xMes = xMes + 1
'            If xMes > 12 Then xMes = 1
'        Next xCol
'    Next xFila
        
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


Sub Cancelar()
    QueHace = 3
    Label2.Caption = "Detalle Plan de Produccion"
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

Sub CambiarEstado(Activado As Boolean)
    Dim Rpta As Integer
    TabOne1.CurrTab = 0
    If Activado = False Then
        Rpta = MsgBox("Esta seguro de desactivar el plan de produccion seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    Else
        Rpta = MsgBox("Esta seguro de activar el plan de produccion seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
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

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then
            MuestraSegundoTab
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        Nuevo
        'CmdAddProd.Visible = True:
        CmdAdd.Visible = True: CmdVerEst.Visible = True
    End If
    
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 8 Then
'        If TabOne2.CurrTab = 0 Then
'            If Fg1.Rows = 1 Then
'                MsgBox "No se ha procesado el plan de produccion para los productos Terminados", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'                Exit Sub
'            End If
'        End If
'        If TabOne2.CurrTab = 1 Then
'            If Fg1.Rows = 1 Then
'                MsgBox "No se ha procesado el plan de produccion para los productos Intermedios", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'                Exit Sub
'            End If
'        End If
'        ExportarExcel
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

Sub MuestraSegundoTab()
    Dim xRst As New ADODB.Recordset
    Dim xSQL As String
    Dim A, B, xMesIni, xAñoTra As Integer
    Dim xMes As String
    
    TxtDesc.Text = RstPlaPro("descripcion")
    TxtFchIni.Valor = Format(RstPlaPro("fchini"), "dd/mm/yyyy")
    TxtFchFin.Valor = Format(RstPlaPro("fchfin"), "dd/mm/yyyy")
    
    xMesIni = RstPlaPro("mesini")
    xAñoTra = RstPlaPro("año")
    
    ' CARGAMOS LOS PRODUCTOS
    xSQL = "TRANSFORM Sum(ges_plaproddet.cantidad) AS SumaDecantidad " _
        & " SELECT ges_plaproddet.idpv, ges_plaproddet.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, Sum(ges_plaproddet.cantidad) AS [Total] " _
        & " FROM (ges_plaproddet LEFT JOIN alm_inventario ON ges_plaproddet.codpro = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        & " Where (((ges_plaproddet.idpv) = " & RstPlaPro("id") & ")) " _
        & " GROUP BY ges_plaproddet.idpv, ges_plaproddet.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev " _
        & " ORDER BY alm_inventario.descripcion PIVOT ges_plaproddet.idmes"

    RST_Busq xRst, xSQL, xCon
    
    Fg1.Rows = 1
    Fg1.Cols = 5
    
    LblNumReg.Caption = Format(xRst.RecordCount, "000")
    
    ' agregamos las columnas de los meses
    For A = xMesIni To 12
        Fg1.Cols = Fg1.Cols + 1
        Fg1.TextMatrix(0, Fg1.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra, "0000")
    Next A
    
    xAñoTra = xAñoTra + 1    ' LE SUMAMOS UN AÑO MAS
    For A = 1 To xMesIni - 1
        Fg1.Cols = Fg1.Cols + 1
        Fg1.TextMatrix(0, Fg1.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra, "0000")
    Next A
    
    xMesIni = RstPlaPro("mesini")
    xAñoTra = RstPlaPro("año")
    
    ' agregamos la columna del total
    Fg1.Cols = Fg1.Cols + 1
    Fg1.TextMatrix(0, Fg1.Cols - 1) = "Programado"
    Fg1.ColWidth(Fg1.Cols - 1) = 1100
    
    If xRst.RecordCount <> 0 Then
        xRst.MoveFirst
        For A = 1 To xRst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = xRst("codpro")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = xRst("idunimed")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = xRst("descripcion")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = xRst("abrev")
            
            ' agregamos los meses
            For B = 5 To 16
                xMes = Trim(Str(NulosN(Mid(Fg1.TextMatrix(0, B), 1, 2))))
                Fg1.TextMatrix(A, B) = Format(xRst(NulosC(xMes)), "0.00")
            Next B

            Fg1.TextMatrix(A, Fg1.Cols - 1) = Format(xRst("Total"), "0.00")
            xRst.MoveNext
            
            If xRst.EOF = True Then Exit For
        Next A
    End If
    
    ' ESCRIBIMOS LOS TOTALES
    Dim xStkIni, xTotPro, xTotal As Double
    'Dim RstTodProd As New Recordset
    
    Fg1.Cols = Fg1.Cols + 4
    
    Fg1.TextMatrix(0, Fg1.Cols - 4) = "Stock Ini"
    Fg1.TextMatrix(0, Fg1.Cols - 3) = "Producido"
    Fg1.TextMatrix(0, Fg1.Cols - 2) = "Total"
    Fg1.TextMatrix(0, Fg1.Cols - 1) = "Diferencia"
    Fg1.ColWidth(Fg1.Cols - 1) = 1100
    
    For A = 1 To Fg1.Rows - 1
        xStkIni = SaldoActual(NulosN(Fg1.TextMatrix(A, 1)), "01/01/" & Format(xAñoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)
        xTotPro = HallarTotalProducido(NulosN(Fg1.TextMatrix(A, 1)), TxtFchIni.Valor)
        
        Fg1.TextMatrix(A, Fg1.Cols - 4) = Format(xStkIni, "0.00")
        Fg1.TextMatrix(A, Fg1.Cols - 3) = Format(xTotPro, "0.00")
        Fg1.TextMatrix(A, Fg1.Cols - 2) = Format(xTotPro + xStkIni, "0.00")
        
        xTotal = ((xTotPro + xStkIni) - NulosN(Fg1.TextMatrix(A, Fg1.Cols - 5)))
        
        If xTotal > 0 Then
            FORMATO_CELDA Fg1, CLng(A), Fg1.Cols - 1, &HFF0000, True, , Format(xTotal, "0.00")
        Else
            FORMATO_CELDA Fg1, CLng(A), Fg1.Cols - 1, &HC0&, True, , Format(xTotal, "0.00")
        End If
    Next A


    
    ' ************************
    ' CARGAMOS LOS INTERMEDIOS
    
    xMesIni = RstPlaPro("mesini")
    xAñoTra = RstPlaPro("año")
    
    xSQL = "TRANSFORM Sum(ges_plaproddet2.cantidad) AS SumaDecantidad" _
        & " SELECT ges_plaproddet2.idpv, ges_plaproddet2.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev, Sum(ges_plaproddet2.cantidad) AS [Total] " _
        & " FROM (ges_plaproddet2 LEFT JOIN alm_inventario ON ges_plaproddet2.codpro = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
        & " Where (((ges_plaproddet2.idpv) = " & RstPlaPro("id") & ")) GROUP BY ges_plaproddet2.idpv, ges_plaproddet2.codpro, alm_inventario.descripcion, alm_inventario.idunimed, mae_unidades.abrev " _
        & " ORDER BY alm_inventario.descripcion PIVOT ges_plaproddet2.idmes"

    RST_Busq xRst, xSQL, xCon
    
    Fg2.Rows = 1
    Fg2.Cols = 6
    ' agregamos las columnas de los meses
    For A = xMesIni To 12
        Fg2.Cols = Fg2.Cols + 1
        Fg2.TextMatrix(0, Fg2.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra, "0000")
    Next A
    
    xAñoTra = xAñoTra + 1    ' LE SUMAMOS UN AÑO MAS
    For A = 1 To xMesIni - 1
        Fg2.Cols = Fg2.Cols + 1
        Fg2.TextMatrix(0, Fg2.Cols - 1) = Format(A, "00") & "-" & Format(xAñoTra, "0000")
    Next A
    
    xMesIni = RstPlaPro("mesini")
    xAñoTra = RstPlaPro("año")
    
    ' agregamos la columna del total
    Fg2.Cols = Fg2.Cols + 1
    Fg2.TextMatrix(0, Fg2.Cols - 1) = "Programado"
    Fg2.ColWidth(Fg2.Cols - 1) = 1100
    
    If xRst.RecordCount <> 0 Then
        xRst.MoveFirst
        For A = 1 To xRst.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = xRst("codpro")
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = xRst("idunimed")
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = ""
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = xRst("descripcion")
            Fg2.TextMatrix(Fg2.Rows - 1, 5) = xRst("abrev")
            
            ' agregamos los meses
            For B = 6 To 17
                xMes = Trim(Str(NulosN(Mid(Fg2.TextMatrix(0, B), 1, 2))))
                Fg2.TextMatrix(A, B) = Format(xRst(NulosC(xMes)), "0.00")
            Next B

            Fg2.TextMatrix(A, Fg2.Cols - 1) = Format(xRst("Total"), "0.00")
            xRst.MoveNext
            
            If xRst.EOF = True Then Exit For
        Next A
    End If
    
    ' ESCRIBIMOS LOS TOTALES
    Fg2.Cols = Fg2.Cols + 4
    
    Fg2.TextMatrix(0, Fg2.Cols - 4) = "Stock Ini"
    Fg2.TextMatrix(0, Fg2.Cols - 3) = "Producido"
    Fg2.TextMatrix(0, Fg2.Cols - 2) = "Total"
    Fg2.TextMatrix(0, Fg2.Cols - 1) = "Diferencia"
    Fg2.ColWidth(Fg2.Cols - 1) = 1100
    
    For A = 1 To Fg2.Rows - 1
        xStkIni = SaldoActual(NulosN(Fg2.TextMatrix(A, 1)), "01/01/" & Format(xAñoTra, "0000"), CDate(TxtFchIni.Valor) - 1, xCon)
        xTotPro = HallarTotalProducido(NulosN(Fg2.TextMatrix(A, 1)), TxtFchIni.Valor)
        
        Fg2.TextMatrix(A, Fg2.Cols - 4) = Format(xStkIni, "0.00")
        Fg2.TextMatrix(A, Fg2.Cols - 3) = Format(xTotPro, "0.00")
        Fg2.TextMatrix(A, Fg2.Cols - 2) = Format(xTotPro + xStkIni, "0.00")
        
        xTotal = ((xTotPro + xStkIni) - NulosN(Fg2.TextMatrix(A, Fg2.Cols - 5)))
        
        If xTotal > 0 Then
            FORMATO_CELDA Fg2, CLng(A), Fg2.Cols - 1, &HFF0000, True, , Format(xTotal, "0.00")
        Else
            FORMATO_CELDA Fg2, CLng(A), Fg2.Cols - 1, &HC0&, True, , Format(xTotal, "0.00")
        End If
    Next A
    
    PintarGrid
    
    Fg1.FrozenCols = 4
    Fg2.FrozenCols = 5
    TabOne2.CurrTab = 0
End Sub

