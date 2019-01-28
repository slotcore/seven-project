VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmManPreven 
   Caption         =   "Mantenimiento - Mantenimiento Preventivo"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13140
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   13140
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   8565
      Left            =   15
      TabIndex        =   11
      Top             =   360
      Width           =   12525
      _cx             =   22093
      _cy             =   15108
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
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
      Caption         =   "   &Consulta   |    &Detalle     "
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
      Begin SizerOneLibCtl.ElasticOne ElasticOne2 
         Height          =   8145
         Left            =   13170
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   375
         Width           =   12435
         _cx             =   21934
         _cy             =   14367
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
         BorderWidth     =   2
         ChildSpacing    =   2
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
         _GridInfo       =   $"FrmManPreven.frx":0000
         Begin SizerOneLibCtl.ElasticOne ElasticOne3 
            Height          =   7410
            Left            =   30
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   360
            Width           =   12375
            _cx             =   21828
            _cy             =   13070
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
            BackColor       =   32896
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   2
            ChildSpacing    =   2
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
            GridRows        =   4
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmManPreven.frx":004F
            Begin SizerOneLibCtl.ElasticOne ElasticOne4 
               Height          =   2715
               Left            =   30
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   4665
               Width           =   12315
               _cx             =   21722
               _cy             =   4789
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
               BorderWidth     =   2
               ChildSpacing    =   2
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
               GridCols        =   3
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"FrmManPreven.frx":00AD
               Begin VSFlex7Ctl.VSFlexGrid Fg4 
                  Height          =   2385
                  Left            =   8175
                  TabIndex        =   10
                  Top             =   300
                  Width           =   4110
                  _cx             =   7250
                  _cy             =   4207
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
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmManPreven.frx":0109
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
               Begin VSFlex7Ctl.VSFlexGrid Fg3 
                  Height          =   2385
                  Left            =   4140
                  TabIndex        =   9
                  Top             =   300
                  Width           =   4005
                  _cx             =   7064
                  _cy             =   4207
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
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmManPreven.frx":01CA
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
                  Height          =   2385
                  Left            =   30
                  TabIndex        =   8
                  Top             =   300
                  Width           =   4080
                  _cx             =   7197
                  _cy             =   4207
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
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmManPreven.frx":0287
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
               Begin VB.Label Label7 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "Herramientas"
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
                  Height          =   240
                  Left            =   8175
                  TabIndex        =   31
                  Top             =   30
                  Width           =   4110
               End
               Begin VB.Label Label6 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Repuestos y Accesorios"
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
                  Height          =   240
                  Left            =   4140
                  TabIndex        =   30
                  Top             =   30
                  Width           =   4005
               End
               Begin VB.Label Label2 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Materiales"
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
                  Height          =   240
                  Left            =   30
                  TabIndex        =   29
                  Top             =   30
                  Width           =   4080
               End
            End
            Begin VSFlex7Ctl.VSFlexGrid Fg1 
               Height          =   1980
               Left            =   30
               TabIndex        =   7
               Top             =   2655
               Width           =   12315
               _cx             =   21722
               _cy             =   3492
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
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManPreven.frx":0344
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
            Begin VB.Frame Frame2 
               BackColor       =   &H00C0E0FF&
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   2325
               Left            =   30
               TabIndex        =   19
               Top             =   30
               Width           =   12315
               Begin VB.TextBox TxtCaracteristicas 
                  BackColor       =   &H00E0E0E0&
                  Height          =   645
                  Left            =   1185
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   4
                  Tag             =   "a"
                  Text            =   "FrmManPreven.frx":041F
                  Top             =   630
                  Width           =   10995
               End
               Begin VB.TextBox TxtMarca 
                  BackColor       =   &H00E0E0E0&
                  Height          =   300
                  Left            =   7155
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   1
                  Tag             =   "a"
                  Text            =   "TxtMarca"
                  Top             =   15
                  Width           =   5040
               End
               Begin VB.TextBox TxtNumSer 
                  BackColor       =   &H00E0E0E0&
                  Height          =   300
                  Left            =   7155
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   3
                  Tag             =   "a"
                  Text            =   "TxtNumSer"
                  Top             =   315
                  Width           =   5040
               End
               Begin VB.TextBox TxtModelo 
                  BackColor       =   &H00E0E0E0&
                  Height          =   300
                  Left            =   1185
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   2
                  Tag             =   "a"
                  Text            =   "TxtModelo"
                  Top             =   315
                  Width           =   5040
               End
               Begin VB.TextBox TxtObserva 
                  Height          =   645
                  Left            =   1185
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   6
                  Tag             =   "a"
                  Text            =   "FrmManPreven.frx":0434
                  Top             =   1650
                  Width           =   10995
               End
               Begin VB.TextBox TxtEquipo 
                  BackColor       =   &H00E0E0E0&
                  Height          =   300
                  Left            =   1185
                  Locked          =   -1  'True
                  TabIndex        =   0
                  Tag             =   "a"
                  Text            =   "TxtEquipo"
                  Top             =   15
                  Width           =   5040
               End
               Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
                  Height          =   300
                  Left            =   1185
                  TabIndex        =   5
                  Top             =   1350
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
                  Locked          =   -1  'True
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Características"
                  Height          =   195
                  Index           =   9
                  Left            =   30
                  TabIndex        =   26
                  Top             =   675
                  Width           =   1065
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Nº Serie"
                  Height          =   195
                  Index           =   5
                  Left            =   6450
                  TabIndex        =   25
                  Top             =   345
                  Width           =   585
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Marca"
                  Height          =   195
                  Index           =   6
                  Left            =   6585
                  TabIndex        =   24
                  Top             =   60
                  Width           =   450
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Modelo"
                  Height          =   195
                  Index           =   0
                  Left            =   30
                  TabIndex        =   23
                  Top             =   345
                  Width           =   525
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Observaciones"
                  Height          =   195
                  Index           =   1
                  Left            =   30
                  TabIndex        =   22
                  Top             =   1710
                  Width           =   1065
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Equipo"
                  Height          =   195
                  Index           =   7
                  Left            =   30
                  TabIndex        =   21
                  Top             =   60
                  Width           =   495
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Fch. Inicio"
                  Height          =   195
                  Index           =   8
                  Left            =   30
                  TabIndex        =   20
                  Top             =   1380
                  Width           =   735
               End
            End
            Begin VB.Label Label1 
               Caption         =   "Tareas a Realizar"
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
               Height          =   240
               Left            =   30
               TabIndex        =   27
               Top             =   2385
               Width           =   12315
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "Detalle Mantenimiento Preventivo"
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
            Height          =   300
            Left            =   30
            TabIndex        =   16
            Top             =   30
            Width           =   12375
         End
      End
      Begin SizerOneLibCtl.ElasticOne ElasticOne1 
         Height          =   8145
         Left            =   45
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   375
         Width           =   12435
         _cx             =   21934
         _cy             =   14367
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
         BorderWidth     =   2
         ChildSpacing    =   2
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
         _GridInfo       =   $"FrmManPreven.frx":0441
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   7740
            Left            =   30
            TabIndex        =   13
            ToolTipText     =   "Click derecho para Aceptar o Rechazar"
            Top             =   375
            Width           =   12375
            _ExtentX        =   21828
            _ExtentY        =   13653
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Codigo"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descripción"
            Columns(1).DataField=   "nombre"
            Columns(1).NumberFormat=   "Short Date"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Marca"
            Columns(2).DataField=   "marca"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Modelo"
            Columns(3).DataField=   "modelo"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Nº Serie"
            Columns(4).DataField=   "serie"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Caracteristica"
            Columns(5).DataField=   "caracteristicas"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Area"
            Columns(6).DataField=   "area"
            Columns(6).NumberFormat=   "0.00"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1244"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1164"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=3704"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3625"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=3307"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3228"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=3519"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3440"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=2910"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2831"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=2778"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2699"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=2223"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2143"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=512"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=66,.parent=13,.alignment=0"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
            _StyleDefs(64)  =   "Named:id=33:Normal"
            _StyleDefs(65)  =   ":id=33,.parent=0"
            _StyleDefs(66)  =   "Named:id=34:Heading"
            _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(68)  =   ":id=34,.wraptext=-1"
            _StyleDefs(69)  =   "Named:id=35:Footing"
            _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(71)  =   "Named:id=36:Selected"
            _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=37:Caption"
            _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(75)  =   "Named:id=38:HighlightRow"
            _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=39:EvenRow"
            _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(79)  =   "Named:id=40:OddRow"
            _StyleDefs(80)  =   ":id=40,.parent=33"
            _StyleDefs(81)  =   "Named:id=41:RecordSelector"
            _StyleDefs(82)  =   ":id=41,.parent=34"
            _StyleDefs(83)  =   "Named:id=42:FilterBar"
            _StyleDefs(84)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0FF&
            Caption         =   "Consulta de Equipos"
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
            Height          =   315
            Index           =   0
            Left            =   30
            TabIndex        =   14
            Top             =   30
            Width           =   12375
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
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPreven.frx":0484
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPreven.frx":09C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPreven.frx":0D5A
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPreven.frx":0EDE
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPreven.frx":1332
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPreven.frx":144A
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPreven.frx":198E
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPreven.frx":1ED2
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPreven.frx":1FE6
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPreven.frx":20FA
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPreven.frx":254E
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPreven.frx":26BA
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   13140
      _ExtentX        =   23178
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
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
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Ficha del Equipo"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menutar 
      Caption         =   "Tareas"
      Visible         =   0   'False
      Begin VB.Menu menutar_01 
         Caption         =   "Agregar Tareas"
      End
      Begin VB.Menu menutar_02 
         Caption         =   "-"
      End
      Begin VB.Menu menutar_03 
         Caption         =   "Eliminar Tareas"
      End
   End
   Begin VB.Menu mat 
      Caption         =   "Materiales"
      Visible         =   0   'False
      Begin VB.Menu mat_01 
         Caption         =   "Agregar Material"
      End
      Begin VB.Menu mat_02 
         Caption         =   "-"
      End
      Begin VB.Menu mat_03 
         Caption         =   "Eliminar Material"
      End
   End
   Begin VB.Menu rep 
      Caption         =   "Repuestos"
      Visible         =   0   'False
      Begin VB.Menu rep_01 
         Caption         =   "Agregar Respuestos y Accesorios"
      End
      Begin VB.Menu rep_02 
         Caption         =   "-"
      End
      Begin VB.Menu rep_03 
         Caption         =   "Eliminar Repuestos y Accesorios"
      End
   End
   Begin VB.Menu herr 
      Caption         =   "Herramientas"
      Visible         =   0   'False
      Begin VB.Menu herr_01 
         Caption         =   "Agregar Herramientas"
      End
      Begin VB.Menu herr_02 
         Caption         =   "-"
      End
      Begin VB.Menu herr_03 
         Caption         =   "Eliminar Herramientas"
      End
   End
End
Attribute VB_Name = "FrmManPreven"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Quehace As Integer
Dim SeEjecuto As Boolean
Dim RstLista As New ADODB.Recordset
Dim Agregando As Boolean
Dim RstTarMat As New ADODB.Recordset
Dim RstTarRep As New ADODB.Recordset
Dim RstTarHer As New ADODB.Recordset
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstLista("id")), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Quehace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    If Col = 1 Then
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
        xForm.SQLCad = "SELECT * FROM man_tareas"
        
        xForm.Titulo = "Buscando Tareas de Mantenimiento"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 1) = xRs("descripcion")
                Fg1.TextMatrix(Fg1.Row, 5) = xRs("id")
            End If
        End If
        Set xForm = Nothing
        Set xRs = Nothing
    End If
    
    If Col = 2 Then
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "cod":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
        xForm.SQLCad = "SELECT * FROM man_frecuencia"
        
        xForm.Titulo = "Buscando Frecuencia de Mantenimiento"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        If xRs.State = 1 Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 2) = xRs("descripcion")
                Fg1.TextMatrix(Fg1.Row, 6) = xRs("id")
            End If
        End If
        Set xForm = Nothing
        Set xRs = Nothing
    End If
End Sub

Private Sub Fg1_EnterCell()
    If Fg1.Col = 1 Or Fg1.Col = 2 Then
        'Fg1.Editable = flexEDNone
    
    End If
    
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Quehace = 3 Then KeyAscii = 0
    
    If KeyAscii = 13 Then Exit Sub
    ' validar los caracteres que se ingresan
    Select Case Col
        Case 1, 2    ' para que no permita el ingreso de caracteres en las columnas donde se seleccione
            KeyAscii = 0
            
        Case 3       ' para el ingreso de la hora
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End Select
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then  ' AGREGAMOS UNA FILA
        If Fg1.TextMatrix(Fg1.Rows - 1, 1) <> "" Then
            
            Fg1.Rows = Fg1.Rows + 1
            Fg1.Select Fg1.Rows - 1, 1
            Fg1_CellButtonClick Fg1.Rows - 1, 1
        End If
    End If

    If KeyCode = 46 Then  ' ELIMINAMOS UNA FILA
        If Fg1.Row < Fg1.FixedRows Then Exit Sub
        RemoverTarea NulosN(Fg1.TextMatrix(Fg1.Row, 5))
        Fg1.RemoveItem Fg1.Row
        Fg2.Rows = 1
        Fg2.Rows = Fg2.Rows + 1
    End If
End Sub

Sub RemoverTarea(Idtarea As Integer)
    Dim A As Integer
    
    ' REMOVEMOS LOS MATERIALES QUE SE LE HAYA ASIGNADO A LA TAREA
    RstTarMat.Filter = adFilterNone
    If RstTarMat.RecordCount <> 0 Then
        RstTarMat.MoveFirst
    End If
    RstTarMat.Filter = "idtar = " & Idtarea
    If RstTarMat.RecordCount <> 0 Then
        RstTarMat.MoveFirst
        For A = 1 To RstTarMat.RecordCount
            RstTarMat.Delete
            
            RstTarMat.MoveNext
            If RstTarMat.EOF = True Then Exit For
        Next A
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu menutar
    End If
End Sub

Private Sub Fg1_RowColChange()
    If Agregando = True Then Exit Sub
    MuestraMaterial NulosN(Fg1.TextMatrix(Fg1.Row, 5))
    MuestraRepuesto NulosN(Fg1.TextMatrix(Fg1.Row, 5))
    MuestraHerramienta NulosN(Fg1.TextMatrix(Fg1.Row, 5))
End Sub

Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Quehace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(3, 4) As String
    
    If Col = 1 Then
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "Uni. Med.":     xCampos(1, 1) = "abrev":          xCampos(1, 2) = "900":     xCampos(1, 3) = "C"
        xCampos(2, 0) = "Codigo":        xCampos(2, 1) = "codpro":         xCampos(2, 2) = "1200":    xCampos(2, 3) = "C"
    
        xForm.SQLCad = "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.tippro, " _
            & " alm_inventario.idunimed FROM alm_inventario LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id WHERE (((alm_inventario.tippro)=7))"

        xForm.Titulo = "Buscando Materiales"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        If xRs.State = 1 Then
            Agregando = True
            If xRs.RecordCount <> 0 Then
                Fg2.TextMatrix(Fg2.Row, 1) = xRs("descripcion")
                Fg2.TextMatrix(Fg2.Row, 2) = xRs("abrev")
                Fg2.TextMatrix(Fg2.Row, 4) = xRs("id")
                Fg2.TextMatrix(Fg2.Row, 5) = xRs("idunimed")
                
                GrabarTmpDatosTareaMaterial NulosN(Fg1.TextMatrix(Fg1.Row, 5)), xRs("id"), xRs("idunimed"), xRs("abrev"), xRs("descripcion"), 0
            End If
            Agregando = False
        End If
        Set xForm = Nothing
        Set xRs = Nothing
    End If
End Sub

Sub GrabarTmpDatosTareaHerramienta(IdTar As Integer, IdMat As Integer, IdUni As Integer, DescAbre As String, DescMat As String, Cantidad As Double)
    RstTarHer.Filter = adFilterNone
    If RstTarHer.RecordCount <> 0 Then
        RstTarHer.MoveFirst
    End If
    RstTarHer.Filter = "idtar = " & IdTar & " and idherr =  " & IdMat & ""
    If RstTarHer.RecordCount = 0 Then
        RstTarHer.AddNew
    End If
    
    RstTarHer("idtar") = IdTar                                      ' ID DEL MATERIAL
    If NulosN(IdMat) <> 0 Then RstTarHer("idherr") = IdMat          ' ID DEL REPUESTO
    If NulosN(IdUni) <> 0 Then RstTarHer("iduni") = IdUni           ' ID DE LA UNIDAD DE MEDIDA
    If NulosC(DescAbre) <> "" Then RstTarHer("descuni") = DescAbre  ' DESCRIPCION DE LA UNIDAD DE MEDIDA
    If NulosC(DescMat) <> "" Then RstTarHer("descrep") = DescMat    ' DESCRIPCION DEL MATERIAL
    If NulosN(Cantidad) <> 0 Then RstTarHer("cantidad") = Cantidad  ' CANTIDAD DEL MATERIAL
End Sub

Sub GrabarTmpDatosTareaRepuesto(IdTar As Integer, IdMat As Integer, IdUni As Integer, DescAbre As String, DescMat As String, Cantidad As Double)
    RstTarRep.Filter = adFilterNone
    If RstTarRep.RecordCount <> 0 Then
        RstTarRep.MoveFirst
    End If
    RstTarRep.Filter = "idtar = " & IdTar & " and Idrep =  " & IdMat & ""
    If RstTarRep.RecordCount = 0 Then
        RstTarRep.AddNew
    End If
    
    RstTarRep("idtar") = IdTar                                      ' ID DEL MATERIAL
    If NulosN(IdMat) <> 0 Then RstTarRep("idrep") = IdMat           ' ID DEL REPUESTO
    If NulosN(IdUni) <> 0 Then RstTarRep("iduni") = IdUni           ' ID DE LA UNIDAD DE MEDIDA
    If NulosC(DescAbre) <> "" Then RstTarRep("descuni") = DescAbre  ' DESCRIPCION DE LA UNIDAD DE MEDIDA
    If NulosC(DescMat) <> "" Then RstTarRep("descrep") = DescMat    ' DESCRIPCION DEL MATERIAL
    If NulosN(Cantidad) <> 0 Then RstTarRep("cantidad") = Cantidad  ' CANTIDAD DEL MATERIAL
End Sub


Sub GrabarTmpDatosTareaMaterial(IdTar As Integer, IdMat As Integer, IdUni As Integer, DescAbre As String, DescMat As String, Cantidad As Double)
    RstTarMat.Filter = adFilterNone
    If RstTarMat.RecordCount <> 0 Then
        RstTarMat.MoveFirst
    End If
    RstTarMat.Filter = "idtar = " & IdTar & " and IdMat =  " & IdMat & ""
    If RstTarMat.RecordCount = 0 Then
        RstTarMat.AddNew
    End If
    
    RstTarMat("idtar") = IdTar                                      ' ID DEL MATERIAL
    If NulosN(IdMat) <> 0 Then RstTarMat("idmat") = IdMat           ' DEL MATERIAL
    If NulosN(IdUni) <> 0 Then RstTarMat("iduni") = IdUni           ' ID DE LA UNIDAD DE MEDIDA
    If NulosC(DescAbre) <> "" Then RstTarMat("descuni") = DescAbre  ' DESCRIPCION DE LA UNIDAD DE MEDIDA
    If NulosC(DescMat) <> "" Then RstTarMat("descmat") = DescMat    ' DESCRIPCION DEL MATERIAL
    If NulosN(Cantidad) <> 0 Then RstTarMat("cantidad") = Cantidad  ' CANTIDAD DEL MATERIAL
End Sub

Sub EliminarTmpDatosTareaHerramienta(IdTar As Integer, IdMat As Integer)
    RstTarHer.Filter = adFilterNone
    If RstTarHer.RecordCount <> 0 Then
        RstTarHer.MoveFirst
    End If
    RstTarHer.Filter = "idtar = " & IdTar & " and idherr =  " & IdMat & ""
    If RstTarHer.RecordCount = 1 Then
        RstTarHer.Delete
    End If
End Sub

Sub EliminarTmpDatosTareaRepuesto(IdTar As Integer, IdMat As Integer)
    RstTarRep.Filter = adFilterNone
    If RstTarRep.RecordCount <> 0 Then
        RstTarRep.MoveFirst
    End If
    RstTarRep.Filter = "idtar = " & IdTar & " and idrep =  " & IdMat & ""
    If RstTarRep.RecordCount = 1 Then
        RstTarRep.Delete
    End If
End Sub

Sub EliminarTmpDatosTareaMaterial(IdTar As Integer, IdMat As Integer)
    RstTarMat.Filter = adFilterNone
    If RstTarMat.RecordCount <> 0 Then
        RstTarMat.MoveFirst
    End If
    RstTarMat.Filter = "idtar = " & IdTar & " and IdMat =  " & IdMat & ""
    If RstTarMat.RecordCount = 1 Then
        RstTarMat.Delete
    End If
End Sub

Private Sub Fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Col = 3 Then
        GrabarTmpDatosTareaMaterial NulosN(Fg1.TextMatrix(Fg1.Row, 5)), NulosN(Fg2.TextMatrix(Fg2.Row, 4)), 0, "", "", NulosN(Fg2.TextMatrix(Fg2.Row, 3))
        Fg2.TextMatrix(Row, Col) = Format(Fg2.TextMatrix(Row, Col), "0.00")
    End If
End Sub

Private Sub Fg2_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Quehace = 3 Then KeyAscii = 0
    
    If KeyAscii = 13 Then Exit Sub
    ' validar los caracteres que se ingresan
    Select Case Col
        Case 1, 2    ' para que no permita el ingreso de caracteres en las columnas donde se seleccione
            KeyAscii = 0
            
        Case 3       ' para el ingreso de la hora
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End Select
End Sub

Private Sub Fg2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then  ' AGREGAMOS UNA FILA
        If Fg2.TextMatrix(Fg2.Rows - 1, 1) <> "" Then
            
            Fg2.Rows = Fg2.Rows + 1
            Fg2.Select Fg2.Rows - 1, 1
            Fg2_CellButtonClick Fg2.Rows - 1, 1
        End If
    End If

    If KeyCode = 46 Then  ' ELIMINAMOS UNA FILA
        EliminarTmpDatosTareaMaterial NulosN(Fg1.TextMatrix(Fg1.Row, 5)), NulosN(Fg2.TextMatrix(Fg2.Row, 4))
        Fg2.RemoveItem Fg2.Row
    End If
End Sub

Private Sub Fg2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mat
    End If
End Sub

Private Sub Fg3_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Quehace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(3, 4) As String
    
    If Col = 1 Then
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "Uni. Med.":     xCampos(1, 1) = "abrev":          xCampos(1, 2) = "900":     xCampos(1, 3) = "C"
        xCampos(2, 0) = "Codigo":        xCampos(2, 1) = "codpro":         xCampos(2, 2) = "1200":    xCampos(2, 3) = "C"
    
        xForm.SQLCad = "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.tippro, " _
            & " alm_inventario.idunimed FROM alm_inventario LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id WHERE (((alm_inventario.tippro)=10))"

        xForm.Titulo = "Buscando Repuestos y Accesorios"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        If xRs.State = 1 Then
            Agregando = True
            If xRs.RecordCount <> 0 Then
                Fg3.TextMatrix(Fg3.Row, 1) = xRs("descripcion")
                Fg3.TextMatrix(Fg3.Row, 2) = xRs("abrev")
                Fg3.TextMatrix(Fg3.Row, 4) = xRs("id")
                Fg3.TextMatrix(Fg3.Row, 5) = xRs("idunimed")
                
                GrabarTmpDatosTareaRepuesto NulosN(Fg1.TextMatrix(Fg1.Row, 5)), xRs("id"), xRs("idunimed"), xRs("abrev"), xRs("descripcion"), 0
            End If
            Agregando = False
        End If
        Set xForm = Nothing
        Set xRs = Nothing
    End If
End Sub

Private Sub Fg3_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Col = 3 Then
        GrabarTmpDatosTareaRepuesto NulosN(Fg1.TextMatrix(Fg1.Row, 5)), NulosN(Fg3.TextMatrix(Fg3.Row, 4)), 0, "", "", NulosN(Fg3.TextMatrix(Fg3.Row, 3))
        Fg3.TextMatrix(Row, Col) = Format(Fg3.TextMatrix(Row, Col), "0.00")
    End If
End Sub

Private Sub Fg3_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Quehace = 3 Then KeyAscii = 0
    
    If KeyAscii = 13 Then Exit Sub
    ' validar los caracteres que se ingresan
    Select Case Col
        Case 1, 2    ' para que no permita el ingreso de caracteres en las columnas donde se seleccione
            KeyAscii = 0
            
        Case 3       ' para el ingreso de la hora
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End Select
End Sub

Private Sub Fg3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then  ' AGREGAMOS UNA FILA
        If Fg3.TextMatrix(Fg3.Rows - 1, 1) <> "" Then
            
            Fg3.Rows = Fg3.Rows + 1
            Fg3.Select Fg3.Rows - 1, 1
            Fg3_CellButtonClick Fg3.Rows - 1, 1
        End If
    End If

    If KeyCode = 46 Then  ' ELIMINAMOS UNA FILA
        EliminarTmpDatosTareaRepuesto NulosN(Fg1.TextMatrix(Fg1.Row, 5)), NulosN(Fg3.TextMatrix(Fg3.Row, 4))
        Fg3.RemoveItem Fg3.Row
    End If
End Sub

Private Sub Fg3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu rep
    End If
End Sub

Private Sub Fg4_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Quehace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(3, 4) As String
    
    If Col = 1 Then
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "Uni. Med.":     xCampos(1, 1) = "abrev":          xCampos(1, 2) = "900":     xCampos(1, 3) = "C"
        xCampos(2, 0) = "Codigo":        xCampos(2, 1) = "codpro":         xCampos(2, 2) = "1200":    xCampos(2, 3) = "C"
    
        xForm.SQLCad = "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.tippro, " _
            & " alm_inventario.idunimed FROM alm_inventario LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id WHERE (((alm_inventario.tippro)=9))"

        xForm.Titulo = "Buscando Herramientas"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        If xRs.State = 1 Then
            Agregando = True
            If xRs.RecordCount <> 0 Then
                Fg4.TextMatrix(Fg4.Row, 1) = xRs("descripcion")
                Fg4.TextMatrix(Fg4.Row, 2) = xRs("abrev")
                Fg4.TextMatrix(Fg4.Row, 4) = xRs("id")
                Fg4.TextMatrix(Fg4.Row, 5) = xRs("idunimed")
                
                GrabarTmpDatosTareaHerramienta NulosN(Fg1.TextMatrix(Fg1.Row, 5)), xRs("id"), xRs("idunimed"), xRs("abrev"), xRs("descripcion"), 0
            End If
            Agregando = False
        End If
        Set xForm = Nothing
        Set xRs = Nothing
    End If
End Sub

Private Sub Fg4_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Col = 3 Then
        GrabarTmpDatosTareaHerramienta NulosN(Fg1.TextMatrix(Fg1.Row, 5)), NulosN(Fg4.TextMatrix(Fg4.Row, 4)), 0, "", "", NulosN(Fg4.TextMatrix(Fg4.Row, 3))
        Fg4.TextMatrix(Row, Col) = Format(Fg4.TextMatrix(Row, Col), "0.00")
    End If
End Sub

Private Sub Fg4_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Quehace = 3 Then KeyAscii = 0
    
    If KeyAscii = 13 Then Exit Sub
    ' validar los caracteres que se ingresan
    Select Case Col
        Case 1, 2    ' para que no permita el ingreso de caracteres en las columnas donde se seleccione
            KeyAscii = 0
            
        Case 3       ' para el ingreso de la hora
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End Select
End Sub

Private Sub Fg4_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then  ' AGREGAMOS UNA FILA
        If Fg4.TextMatrix(Fg4.Rows - 1, 1) <> "" Then
            
            Fg4.Rows = Fg4.Rows + 1
            Fg4.Select Fg4.Rows - 1, 1
            Fg4_CellButtonClick Fg4.Rows - 1, 1
        End If
    End If

    If KeyCode = 46 Then  ' ELIMINAMOS UNA FILA
        EliminarTmpDatosTareaHerramienta NulosN(Fg1.TextMatrix(Fg1.Row, 5)), NulosN(Fg4.TextMatrix(Fg4.Row, 4))
        Fg4.RemoveItem Fg4.Row
    End If
End Sub

Private Sub Fg4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu herr
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        
        RST_Busq RstLista, "SELECT man_equipos.*, pla_area.descripcion AS area, man_preventivo.observaciones AS obsmanpre, man_preventivo.fchini" _
            & " FROM (man_equipos LEFT JOIN pla_area ON man_equipos.idarea = pla_area.id) LEFT JOIN man_preventivo ON man_equipos.id = man_preventivo.idequipo " _
            & " ORDER BY man_equipos.nombre", xCon

        Set Dg1.DataSource = RstLista
        If RstLista.RecordCount = 0 Then
            MsgBox "No se hay equipos registrados, vaya a mantenimiento de equipos y registre al menos un equipo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            RstLista.Close
            Set RstLista = Nothing
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    Quehace = 3
    TabOne1.CurrTab = 0
    ConfiguraForm
    TabOne1.Left = 15
    TabOne1.Top = 360
    SeEjecuto = False
End Sub

Sub ActivaTool()
    Dim A As Integer
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Bloquea()
    'TxtEquipo.Locked = Not TxtEquipo.Locked
    TxtFchIni.Locked = Not TxtFchIni.Locked
    TxtObserva.Locked = Not TxtObserva.Locked
End Sub

Function MuestraSegundoTab() As Boolean
    MuestraSegundoTab = False
    Dim RstPre As New ADODB.Recordset
    Dim RstTar As New ADODB.Recordset
    Dim A, Rpta As Integer
    
    If Quehace = 3 Then
        RST_Busq RstPre, "SELECT * FROM man_preventivo where idequipo = " & RstLista("id") & "", xCon
        If RstPre.RecordCount = 0 Then
            MsgBox "No se ha definido el mantenimiento preventivo para este equipo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TabOne1.CurrTab = 0
            RstPre.Close
            Set RstPre = Nothing
            MuestraSegundoTab = True
            Exit Function
        End If
        
        RstPre.Close
        Set RstPre = Nothing
    End If
    
    Fg1.Rows = 1
    Fg2.Rows = 1
    Fg3.Rows = 1
    Fg4.Rows = 1
    
    TxtEquipo.Text = RstLista("nombre")
    TxtMarca.Text = NulosC(RstLista("marca"))
    TxtModelo.Text = NulosC(RstLista("modelo"))
    TxtNumSer.Text = NulosC(RstLista("serie"))
    TxtCaracteristicas.Text = NulosC(RstLista("caracteristicas"))
    
    TxtFchIni.Valor = NulosC(RstLista("fchini"))
    TxtObserva.Text = NulosC(RstLista("obsmanpre"))
    
    RST_Busq RstTar, "SELECT man_tareas.descripcion AS desctar, man_frecuencia.descripcion AS descfre, man_preventivotarea.* " _
        & " FROM man_tareas INNER JOIN (man_frecuencia INNER JOIN man_preventivotarea ON man_frecuencia.id = man_preventivotarea.idfre) " _
        & " ON man_tareas.id = man_preventivotarea.idtar WHERE (((man_preventivotarea.idequipo)=" & RstLista("id") & "))", xCon

    If RstTar.RecordCount <> 0 Then
        Fg1.Rows = 1
        RstTar.MoveFirst
        Agregando = True
        For A = 1 To RstTar.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = RstTar("desctar")
            Fg1.TextMatrix(A, 2) = RstTar("descfre")
            If IsNull(RstTar("tiempo")) = False Then
                Fg1.TextMatrix(A, 3) = Format(RstTar("tiempo"), "hh:ss")
            Else
                Fg1.TextMatrix(A, 3) = ""
            End If
            Fg1.TextMatrix(A, 4) = RstTar("observa")
            Fg1.TextMatrix(A, 5) = RstTar("idtar")
            Fg1.TextMatrix(A, 6) = RstTar("idfre")
            RstTar.MoveNext
            If RstTar.EOF = True Then
                Exit For
            End If
        Next A
        Agregando = False
    Else
        Fg1.Rows = Fg1.Rows + 1
    End If
    
    ' CARGAMOS LOS MATERIALES A UTILIZAR POR TAREA
    RST_Busq RstTarMat, "SELECT man_preventivotareamat.*, alm_inventario.descripcion AS descmat, mae_unidades.abrev AS descuni " _
        & " FROM (man_preventivotareamat LEFT JOIN alm_inventario ON man_preventivotareamat.idmat = alm_inventario.id) LEFT JOIN mae_unidades " _
        & " ON man_preventivotareamat.iduni = mae_unidades.id WHERE (((man_preventivotareamat.idequipo) = " & RstLista("id") & "))", xCon

    If RstTarMat.RecordCount <> 0 Then
        MuestraMaterial NulosN(Fg1.TextMatrix(1, 5))
    Else
        Fg2.Rows = Fg2.Rows + 1
    End If
    RstTarMat.ActiveConnection = Nothing
    
    
    ' CARGAMOS LOS REPUESTOS Y ACCESORIOS A UTILIZAR POR TAREA
    RST_Busq RstTarRep, "SELECT man_preventivotarearep.*, alm_inventario.descripcion AS descrep, mae_unidades.abrev AS descuni" _
        & " FROM (man_preventivotarearep LEFT JOIN alm_inventario ON man_preventivotarearep.idrep = alm_inventario.id) LEFT JOIN mae_unidades " _
        & " ON alm_inventario.idunimed = mae_unidades.id WHERE (((man_preventivotarearep.idequipo) = " & RstLista("id") & "))", xCon
    
    If RstTarRep.RecordCount <> 0 Then
        MuestraRepuesto NulosN(Fg1.TextMatrix(1, 5))
    Else
        Fg3.Rows = Fg3.Rows + 1
    End If
    RstTarRep.ActiveConnection = Nothing


    ' CARGAMOS LAS HERRAMIENTAS A UTILIZAR POR LA TAREA
    RST_Busq RstTarHer, "SELECT man_preventivotareaherr.*, alm_inventario.descripcion AS descrep, mae_unidades.abrev AS descuni" _
        & " FROM (man_preventivotareaherr LEFT JOIN alm_inventario ON man_preventivotareaherr.idherr = alm_inventario.id) LEFT JOIN mae_unidades " _
        & " ON alm_inventario.idunimed = mae_unidades.id WHERE (((man_preventivotareaherr.idequipo)=" & RstLista("id") & "))", xCon
    
    If RstTarHer.RecordCount <> 0 Then
        MuestraHerramienta NulosN(Fg1.TextMatrix(1, 5))
    Else
        Fg4.Rows = Fg4.Rows + 1
    End If
    RstTarHer.ActiveConnection = Nothing
End Function

Sub MuestraRepuesto(Idtarea As Integer)
    Dim A As Integer
    RstTarRep.Filter = adFilterNone
    RstTarRep.Filter = "idtar = " & Idtarea & ""
    Fg3.Rows = 1
    If RstTarRep.RecordCount <> 0 Then
        RstTarRep.MoveFirst
        Agregando = True
        For A = 1 To RstTarRep.RecordCount
            Fg3.Rows = Fg3.Rows + 1
            Fg3.TextMatrix(A, 1) = NulosC(RstTarRep("descrep"))
            Fg3.TextMatrix(A, 2) = NulosC(RstTarRep("descuni"))
            Fg3.TextMatrix(A, 3) = Format(NulosN(RstTarRep("cantidad")), "0.00")
            Fg3.TextMatrix(A, 4) = NulosN(RstTarRep("idrep"))
            Fg3.TextMatrix(A, 5) = NulosN(RstTarRep("iduni"))
            RstTarRep.MoveNext
            If RstTarRep.EOF = True Then Exit For
        Next A
        Agregando = False
    Else
        Fg3.Rows = Fg3.Rows + 1
    End If
End Sub

Sub MuestraHerramienta(Idtarea As Integer)
    Dim A As Integer
    RstTarHer.Filter = adFilterNone
    RstTarHer.Filter = "idtar = " & Idtarea & ""
    Fg4.Rows = 1
    If RstTarHer.RecordCount <> 0 Then
        RstTarHer.MoveFirst
        Agregando = True
        For A = 1 To RstTarHer.RecordCount
            Fg4.Rows = Fg4.Rows + 1
            Fg4.TextMatrix(A, 1) = NulosC(RstTarHer("descrep"))
            Fg4.TextMatrix(A, 2) = NulosC(RstTarHer("descuni"))
            Fg4.TextMatrix(A, 3) = Format(NulosN(RstTarHer("cantidad")), "0.00")
            Fg4.TextMatrix(A, 4) = NulosN(RstTarHer("idherr"))
            Fg4.TextMatrix(A, 5) = NulosN(RstTarHer("iduni"))
            RstTarHer.MoveNext
            If RstTarHer.EOF = True Then Exit For
        Next A
        Agregando = False
    Else
        Fg4.Rows = Fg4.Rows + 1
    End If
End Sub

Sub MuestraMaterial(Idtarea As Integer)
    Dim A As Integer
    RstTarMat.Filter = adFilterNone
    RstTarMat.Filter = "idtar = " & Idtarea & ""
    Fg2.Rows = 1
    If RstTarMat.RecordCount <> 0 Then
        RstTarMat.MoveFirst
        Agregando = True
        For A = 1 To RstTarMat.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(A, 1) = NulosC(RstTarMat("descmat"))
            Fg2.TextMatrix(A, 2) = NulosC(RstTarMat("descuni"))
            Fg2.TextMatrix(A, 3) = Format(NulosN(RstTarMat("cantidad")), "0.00")
            Fg2.TextMatrix(A, 4) = NulosN(RstTarMat("idmat"))
            Fg2.TextMatrix(A, 5) = NulosN(RstTarMat("iduni"))
            RstTarMat.MoveNext
            If RstTarMat.EOF = True Then Exit For
        Next A
        Agregando = False
    Else
        Fg2.Rows = Fg2.Rows + 1
    End If
End Sub

Sub Modificar()
    Quehace = 2
    xHorIni = Time
    ActivaTool
    TabOne1.CurrTab = 1
    Label5.Caption = "Definiendo Mantenimiento Preventivo"
    TabOne1.TabEnabled(0) = False
    
    Fg1.Editable = flexEDKbdMouse
    Fg2.Editable = flexEDKbdMouse
    Fg3.Editable = flexEDKbdMouse
    Fg4.Editable = flexEDKbdMouse
    
    Fg1.SelectionMode = flexSelectionFree
    Fg2.SelectionMode = flexSelectionFree
    Fg3.SelectionMode = flexSelectionFree
    Fg4.SelectionMode = flexSelectionFree
    
    Fg1.ColComboList(1) = "|..."
    Fg1.ColComboList(2) = "|..."
    
    Fg2.ColComboList(1) = "|..."
    Fg3.ColComboList(1) = "|..."
    Fg4.ColComboList(1) = "|..."
    Bloquea
End Sub

Sub Blanquea()
    TxtEquipo.Text = ""
    TxtMarca.Text = ""
    TxtModelo.Text = ""
    TxtNumSer.Text = ""
    TxtCaracteristicas.Text = ""
    
    TxtFchIni.Valor = ""
    TxtObserva.Text = ""
    Fg1.Rows = 1
    Fg2.Rows = 1
    Fg3.Rows = 1
    Fg4.Rows = 1
End Sub

Sub ConfiguraForm()
    ElasticOne3.BackColor = &H8000000F
    
    Label5.BackColor = &H8000000F
    Label2.BackColor = &H8000000F
    Label6.BackColor = &H8000000F
    Label7.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    
    Label4(0).BackColor = &H8000000F
    Fg1.ColWidth(5) = 0
    Fg1.ColWidth(6) = 0

    Fg2.ColWidth(4) = 0
    Fg2.ColWidth(5) = 0
    
    Fg3.ColWidth(4) = 0
    Fg3.ColWidth(5) = 0

    Fg4.ColWidth(4) = 0
    Fg4.ColWidth(5) = 0
    
    DefinirAnchoColGrid
    
    Fg2.AllowUserResizing = flexResizeColumns

    Fg1.SelectionMode = flexSelectionByRow
    Fg2.SelectionMode = flexSelectionByRow
    Fg3.SelectionMode = flexSelectionByRow
    Fg4.SelectionMode = flexSelectionByRow
    
    Fg1.BackColorSel = &H40&
    Fg2.BackColorSel = &H40&
    Fg3.BackColorSel = &H40&
    Fg4.BackColorSel = &H40&
End Sub

Private Sub Form_Resize()
    TabOne1.Width = Me.Width - 130
    TabOne1.Height = (Me.Height - 760)
    
    DefinirAnchoColGrid
End Sub

Sub DefinirAnchoColGrid()
    Fg2.ColWidth(1) = ((Fg2.Width - 420) - 1500)
    Fg3.ColWidth(1) = ((Fg3.Width - 420) - 1500)
    Fg4.ColWidth(1) = ((Fg4.Width - 420) - 1500)
End Sub

Private Sub herr_01_Click()
    If Fg4.TextMatrix(Fg4.Rows - 1, 1) <> "" Then
        Fg4.Rows = Fg4.Rows + 1
    End If
End Sub

Private Sub herr_03_Click()
    EliminarTmpDatosTareaRepuesto NulosN(Fg1.TextMatrix(Fg1.Row, 5)), NulosN(Fg4.TextMatrix(Fg4.Row, 4))
    Fg4.RemoveItem Fg4.Row
End Sub

Private Sub mat_01_Click()
    If Fg2.TextMatrix(Fg2.Rows - 1, 1) <> "" Then
        Fg2.Rows = Fg2.Rows + 1
    End If
End Sub

Private Sub mat_03_Click()
    EliminarTmpDatosTareaMaterial NulosN(Fg1.TextMatrix(Fg1.Row, 5)), NulosN(Fg2.TextMatrix(Fg2.Row, 4))
    Fg2.RemoveItem Fg2.Row
End Sub

Private Sub menutar_01_Click()
    If Fg1.TextMatrix(Fg1.Rows - 1, 1) <> "" Then
        Fg1.Rows = Fg1.Rows + 1
    End If
End Sub

Private Sub menutar_03_Click()
    If Fg1.Rows <> 1 Then
        Fg1.RemoveItem Fg1.Row
    End If
End Sub

Private Sub rep_01_Click()
    If Fg3.TextMatrix(Fg3.Rows - 1, 1) <> "" Then
        Fg3.Rows = Fg3.Rows + 1
    End If
End Sub

Private Sub rep_03_Click()
    EliminarTmpDatosTareaRepuesto NulosN(Fg1.TextMatrix(Fg1.Row, 5)), NulosN(Fg3.TextMatrix(Fg3.Row, 4))
    Fg3.RemoveItem Fg3.Row
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        Cancel = MuestraSegundoTab
    End If
End Sub

Function Grabar() As Boolean
    Dim xId As Double
    Dim xCamposPre(2, 5) As String
    Dim xCamposTar(4, 5) As String
    Dim xCamposTarMat(4, 5) As String
    Dim xCamposTarRep(4, 5) As String
    Dim xCamposTarHer(4, 5) As String
    Dim A As Integer
    
On Error GoTo LaCague

    xCon.BeginTrans
    
    If Quehace = 1 Then
        xId = HallaCodigoTabla("man_equipos", xCon, "id")
    Else
        xId = RstLista("id")
        xCon.Execute "DELETE * FROM  man_preventivo WHERE idequipo = " & xId & ""
        xCon.Execute "DELETE * FROM  man_preventivotarea WHERE idequipo = " & xId & ""
        xCon.Execute "DELETE * FROM  man_preventivotareamat WHERE idequipo = " & xId & ""
        xCon.Execute "DELETE * FROM  man_preventivotarearep WHERE idequipo = " & xId & ""
        xCon.Execute "DELETE * FROM  man_preventivotareaherr WHERE idequipo = " & xId & ""
    End If
   
    'Columna    | Descripcion
    '------------------------
    '0          | campo
    '1          | Valor
    '2          | requerido
    '3          | tipo
    '4          | msj error
    
    ' GRABAMOS DATOS EN LA TABLA man_preventivo
    xCamposPre(0, 0) = "idequipo":      xCamposPre(0, 1) = Str(xId):           xCamposPre(0, 2) = "S":  xCamposPre(0, 3) = "N":     xCamposPre(0, 4) = "":                                        xCamposPre(0, 5) = "S"
    xCamposPre(1, 0) = "fchini":        xCamposPre(1, 1) = TxtFchIni.Valor:    xCamposPre(1, 2) = "S":  xCamposPre(1, 3) = "F":     xCamposPre(1, 4) = "No ha especificado la fecha de inicio":   xCamposPre(1, 5) = ""
    xCamposPre(2, 0) = "observaciones": xCamposPre(2, 1) = TxtObserva.Text:    xCamposPre(2, 2) = "N":  xCamposPre(2, 3) = "C":     xCamposPre(2, 4) = "":                                        xCamposPre(2, 5) = ""
    
    If EscribirNuevoRegistro(xCamposPre, "man_preventivo", xCon) = False Then
        xCon.RollbackTrans
        Exit Function
    End If
    
    ' GRABAMOS DATOS EN LA TABLA man_preventivotarea
    For A = 1 To Fg1.Rows - 1
        xCamposTar(0, 0) = "idequipo":  xCamposTar(0, 1) = Str(xId):               xCamposTar(0, 2) = "S":    xCamposTar(0, 3) = "N":     xCamposTar(0, 4) = "":                                    xCamposTar(0, 5) = "S"
        xCamposTar(1, 0) = "idtar":     xCamposTar(1, 1) = Fg1.TextMatrix(A, 5):   xCamposTar(1, 2) = "S":    xCamposTar(1, 3) = "N":     xCamposTar(1, 4) = "No ha especificado la tarea":         xCamposTar(1, 5) = ""
        xCamposTar(2, 0) = "idfre":     xCamposTar(2, 1) = Fg1.TextMatrix(A, 6):   xCamposTar(2, 2) = "S":    xCamposTar(2, 3) = "N":     xCamposTar(2, 4) = "No ha especificado la frecuencia":    xCamposTar(2, 5) = ""
        xCamposTar(3, 0) = "observa":   xCamposTar(3, 1) = Fg1.TextMatrix(A, 4):   xCamposTar(3, 2) = "N":    xCamposTar(3, 3) = "C":     xCamposTar(3, 4) = "":                                    xCamposTar(3, 5) = ""
        xCamposTar(4, 0) = "tiempo":    xCamposTar(4, 1) = Fg1.TextMatrix(A, 3):   xCamposTar(4, 2) = "S":    xCamposTar(4, 3) = "F":     xCamposTar(4, 4) = "":                                    xCamposTar(4, 5) = ""
        
        If EscribirNuevoRegistro(xCamposTar, "man_preventivotarea", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
    Next A
    
    ' GRABAMOS DATOS EN LA TABLA man_preventivotareamat
    RstTarMat.Filter = adFilterNone
    RstTarMat.MoveFirst
    For A = 1 To RstTarMat.RecordCount
        xCamposTarMat(0, 0) = "idtar":       xCamposTarMat(0, 1) = RstTarMat("idtar"):       xCamposTarMat(0, 2) = "S":    xCamposTarMat(0, 3) = "N":     xCamposTarMat(0, 4) = "No ha especificado la tarea":                           xCamposTarMat(0, 5) = "S"
        xCamposTarMat(1, 0) = "idmat":       xCamposTarMat(1, 1) = RstTarMat("idmat"):       xCamposTarMat(1, 2) = "S":    xCamposTarMat(1, 3) = "N":     xCamposTarMat(1, 4) = "No ha especificado el material a utilizar":             xCamposTarMat(1, 5) = ""
        xCamposTarMat(2, 0) = "iduni":       xCamposTarMat(2, 1) = RstTarMat("iduni"):       xCamposTarMat(2, 2) = "S":    xCamposTarMat(2, 3) = "N":     xCamposTarMat(2, 4) = "No ha especificado la unidad de medida":                xCamposTarMat(2, 5) = ""
        xCamposTarMat(3, 0) = "cantidad":    xCamposTarMat(3, 1) = RstTarMat("cantidad"):    xCamposTarMat(3, 2) = "S":    xCamposTarMat(3, 3) = "N":     xCamposTarMat(3, 4) = "No ha especificado la cantidad a utilizar":             xCamposTarMat(3, 5) = ""
        xCamposTarMat(4, 0) = "idequipo":    xCamposTarMat(4, 1) = xId:                      xCamposTarMat(4, 2) = "S":    xCamposTarMat(4, 3) = "N":     xCamposTarMat(4, 4) = "No ha especificado el codigo del equipo":               xCamposTarMat(4, 5) = ""
        
        If EscribirNuevoRegistro(xCamposTarMat, "man_preventivotareamat", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
        RstTarMat.MoveNext
        If RstTarMat.EOF = True Then Exit For
    Next A
    
    ' GRABAMOS LOS DATOS EN LA TABLA man_preventivotarearep
    RstTarRep.Filter = adFilterNone
    RstTarRep.MoveFirst
    For A = 1 To RstTarRep.RecordCount
        xCamposTarRep(0, 0) = "idtar":       xCamposTarRep(0, 1) = RstTarRep("idtar"):       xCamposTarRep(0, 2) = "S":    xCamposTarRep(0, 3) = "N":     xCamposTarRep(0, 4) = "No ha especificado la tarea":                           xCamposTarRep(0, 5) = "S"
        xCamposTarRep(1, 0) = "idrep":       xCamposTarRep(1, 1) = RstTarRep("idrep"):       xCamposTarRep(1, 2) = "S":    xCamposTarRep(1, 3) = "N":     xCamposTarRep(1, 4) = "No ha especificado el material a utilizar":             xCamposTarRep(1, 5) = ""
        xCamposTarRep(2, 0) = "iduni":       xCamposTarRep(2, 1) = RstTarRep("iduni"):       xCamposTarRep(2, 2) = "S":    xCamposTarRep(2, 3) = "N":     xCamposTarRep(2, 4) = "No ha especificado la unidad de medida":                xCamposTarRep(2, 5) = ""
        xCamposTarRep(3, 0) = "cantidad":    xCamposTarRep(3, 1) = RstTarRep("cantidad"):    xCamposTarRep(3, 2) = "S":    xCamposTarRep(3, 3) = "N":     xCamposTarRep(3, 4) = "No ha especificado la cantidad a utilizar":             xCamposTarRep(3, 5) = ""
        xCamposTarRep(4, 0) = "idequipo":    xCamposTarRep(4, 1) = xId:                      xCamposTarRep(4, 2) = "S":    xCamposTarRep(4, 3) = "N":     xCamposTarRep(4, 4) = "No ha especificado el codigo del equipo":               xCamposTarRep(4, 5) = ""
        
        If EscribirNuevoRegistro(xCamposTarRep, "man_preventivotarearep", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
        RstTarRep.MoveNext
        If RstTarRep.EOF = True Then Exit For
    Next A
    
    
    ' GRABAMOS LOS DATOS EN LA TABLA man_preventivotareaherr
    RstTarHer.Filter = adFilterNone
    RstTarHer.MoveFirst
    For A = 1 To RstTarHer.RecordCount
        xCamposTarHer(0, 0) = "idtar":       xCamposTarHer(0, 1) = RstTarHer("idtar"):       xCamposTarHer(0, 2) = "S":    xCamposTarHer(0, 3) = "N":     xCamposTarHer(0, 4) = "No ha especificado la tarea":                 xCamposTarHer(0, 5) = "S"
        xCamposTarHer(1, 0) = "idherr":      xCamposTarHer(1, 1) = RstTarHer("idherr"):       xCamposTarHer(1, 2) = "S":    xCamposTarHer(1, 3) = "N":     xCamposTarHer(1, 4) = "No ha especificado el material a utilizar":   xCamposTarHer(1, 5) = ""
        xCamposTarHer(2, 0) = "iduni":       xCamposTarHer(2, 1) = RstTarHer("iduni"):       xCamposTarHer(2, 2) = "S":    xCamposTarHer(2, 3) = "N":     xCamposTarHer(2, 4) = "No ha especificado la unidad de medida":      xCamposTarHer(2, 5) = ""
        xCamposTarHer(3, 0) = "cantidad":    xCamposTarHer(3, 1) = RstTarHer("cantidad"):    xCamposTarRep(3, 2) = "S":    xCamposTarHer(3, 3) = "N":     xCamposTarHer(3, 4) = "No ha especificado la cantidad a utilizar":   xCamposTarHer(3, 5) = ""
        xCamposTarHer(4, 0) = "idequipo":    xCamposTarHer(4, 1) = xId:                      xCamposTarHer(4, 2) = "S":    xCamposTarHer(4, 3) = "N":     xCamposTarHer(4, 4) = "No ha especificado el codigo del equipo":     xCamposTarHer(4, 5) = ""
        
        If EscribirNuevoRegistro(xCamposTarHer, "man_preventivotareaherr", xCon) = False Then
            xCon.RollbackTrans
            Exit Function
        End If
        RstTarHer.MoveNext
        If RstTarHer.EOF = True Then Exit For
    Next A
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, Quehace, xHorIni, Time, Date, xCon, xId

    xCon.CommitTrans
    MsgBox "Se definio con exito el mantenimiento preventivo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

Sub Cancelar()
    Quehace = 3
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    
    Fg1.SelectionMode = flexSelectionByRow
    Fg2.SelectionMode = flexSelectionByRow
    Fg3.SelectionMode = flexSelectionByRow
    Fg4.SelectionMode = flexSelectionByRow
    
    Bloquea
    ActivaTool
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 2 Then Modificar
'    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstLista.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 14 Then
        RstLista.Close
        Set RstLista = Nothing
        Unload Me
    End If
End Sub

Private Sub TxtCaracteristicas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtEquipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtMarca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtModelo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtObserva_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub
