VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManPersTarea 
   Caption         =   "Producción - Personal de Producción por Tarea"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
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
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPersTarea.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPersTarea.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPersTarea.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPersTarea.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPersTarea.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPersTarea.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPersTarea.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPersTarea.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPersTarea.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPersTarea.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPersTarea.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPersTarea.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPersTarea.frx":277E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPersTarea.frx":2B10
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPersTarea.frx":2E2A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6915
      Left            =   0
      TabIndex        =   1
      Top             =   375
      Width           =   11175
      _cx             =   19711
      _cy             =   12197
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
         Height          =   6495
         Left            =   11820
         TabIndex        =   5
         Top             =   375
         Width           =   11085
         Begin VB.Frame frmTarea 
            Caption         =   "[ Detalle de Tareas ]"
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
            Height          =   1635
            Left            =   60
            TabIndex        =   18
            Top             =   1290
            Width           =   9150
            Begin VSFlex7Ctl.VSFlexGrid Fg 
               Height          =   1305
               Index           =   0
               Left            =   120
               TabIndex        =   19
               Top             =   255
               Width           =   8925
               _cx             =   15743
               _cy             =   2302
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
               Rows            =   4
               Cols            =   6
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManPersTarea.frx":3144
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
         Begin VB.Frame frmPersonal 
            Caption         =   "[ Detalle de Personal ]"
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
            Height          =   3435
            Left            =   60
            TabIndex        =   13
            Top             =   3000
            Width           =   10920
            Begin VSFlex7Ctl.VSFlexGrid Fg 
               Height          =   3045
               Index           =   1
               Left            =   120
               TabIndex        =   14
               Top             =   240
               Width           =   8910
               _cx             =   15716
               _cy             =   5371
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
               Rows            =   4
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManPersTarea.frx":31FD
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
            Begin VB.Frame frmBotones 
               Height          =   3135
               Left            =   9120
               TabIndex        =   15
               Top             =   150
               Width           =   1695
               Begin VB.CommandButton Cmd 
                  Caption         =   "Eliminar &Todos"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   4
                  Left            =   150
                  TabIndex        =   22
                  ToolTipText     =   "Procesa los valores de Linea"
                  Top             =   1620
                  Width           =   1400
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "&Eliminar"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   3
                  Left            =   150
                  TabIndex        =   21
                  ToolTipText     =   "Procesa los valores de Linea"
                  Top             =   1260
                  Width           =   1400
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "&Agregar"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   1
                  Left            =   150
                  TabIndex        =   17
                  ToolTipText     =   "Establece como Principal la Linea Actual "
                  Top             =   150
                  Width           =   1400
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "&Seleccionar"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   2
                  Left            =   150
                  TabIndex        =   16
                  ToolTipText     =   "Carga los valores correspondientes en la Receta"
                  Top             =   510
                  Width           =   1400
               End
            End
         End
         Begin VB.Frame FrmReceta 
            Caption         =   "[ Receta ]"
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
            Height          =   765
            Left            =   60
            TabIndex        =   8
            Top             =   450
            Width           =   10950
            Begin VB.CommandButton Cmd 
               Enabled         =   0   'False
               Height          =   240
               Index           =   0
               Left            =   2520
               Picture         =   "FrmManPersTarea.frx":3302
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   330
               Width           =   225
            End
            Begin VB.TextBox TxtCodRec 
               Height          =   300
               Left            =   1170
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   0
               Text            =   "TxtCodRec"
               Top             =   300
               Width           =   1605
            End
            Begin VB.Label lblIdRec 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "lblIdRec"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   10080
               TabIndex        =   20
               Top             =   360
               Visible         =   0   'False
               Width           =   585
            End
            Begin VB.Label lbl_cb_capt 
               AutoSize        =   -1  'True
               Caption         =   "Decripción"
               Height          =   195
               Index           =   0
               Left            =   150
               TabIndex        =   10
               Top             =   360
               Width           =   765
            End
            Begin VB.Label LblDesRec 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblDesRec"
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
               Left            =   2790
               TabIndex        =   11
               Top             =   315
               Width           =   8010
            End
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   75
            TabIndex        =   12
            Top             =   3435
            Width           =   6330
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Producto"
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
            Left            =   60
            TabIndex        =   6
            Top             =   100
            Width           =   10980
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6495
         Left            =   45
         TabIndex        =   2
         Top             =   375
         Width           =   11085
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   5925
            Left            =   30
            TabIndex        =   3
            Top             =   495
            Width           =   11010
            _ExtentX        =   19420
            _ExtentY        =   10451
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
            Columns(1).Caption=   "Descripción"
            Columns(1).DataField=   "descripcion"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Receta"
            Columns(2).DataField=   "codrec"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=10451"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=10372"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=3228"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=3149"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Named:id=33:Normal"
            _StyleDefs(49)  =   ":id=33,.parent=0"
            _StyleDefs(50)  =   "Named:id=34:Heading"
            _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(52)  =   ":id=34,.wraptext=-1"
            _StyleDefs(53)  =   "Named:id=35:Footing"
            _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(55)  =   "Named:id=36:Selected"
            _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=37:Caption"
            _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(59)  =   "Named:id=38:HighlightRow"
            _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(61)  =   "Named:id=39:EvenRow"
            _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(63)  =   "Named:id=40:OddRow"
            _StyleDefs(64)  =   ":id=40,.parent=33"
            _StyleDefs(65)  =   "Named:id=41:RecordSelector"
            _StyleDefs(66)  =   ":id=41,.parent=34"
            _StyleDefs(67)  =   "Named:id=42:FilterBar"
            _StyleDefs(68)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Producto"
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
            Left            =   30
            TabIndex        =   4
            Top             =   100
            Width           =   11010
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   1005
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Actualizar Costo en Lote"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Equivalencia de Costo en Horas"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmManPersTarea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMMANCOSTO.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO QUE PERMITE ASIGNAR COSTO A LAS TAREAS A CADA RECETA
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 05/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstFrm As New ADODB.Recordset       ' RECORDSET QUE ALAMCENARA LOS PRODCUTOS DISPONIBLES
Dim QueHace As Integer                  ' ESPECIFICA EN QUE MODO SE ENCUENTRA EL FORMULARIO
Dim Agregando As Boolean                ' INDICA QUE SE ESTAN AGREGANDO FILAS A UN CONTROL FLEXGRID
Dim SeEjecuto As Boolean                ' CONTROLA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim mIdRegistro&                        ' identificador del registro
Dim fOrdenLista As Boolean              ' especfica el orden de la lista de la consulta
Dim RstValores As New ADODB.Recordset
Dim IdMenuActivo As Integer             'INDICA EL CODIGO DEL MENU ACTIVO
Dim xHorIni As Date                     'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim cSQL As String
Dim ELIMINARTODO_ As Boolean

Dim COLUMNADETTAREA_ As Integer
Dim COLUMNATOTPERS_ As Integer
Dim COLUMNAPERSSEL_ As Integer
Dim COLUMNAPERSACT_ As Integer
Dim COLUMNAIDTAR_ As Integer

Dim COLUMNAORDEN_ As Integer
Dim COLUMNADETPERS_ As Integer
Dim COLUMNANUMDOC_ As Integer
Dim COLUMNAACTIVO_ As Integer
Dim COLUMNAIDPERS_ As Integer
Dim COLUMNAFCHINGR_ As Integer
Dim COLUMNAFCHCESE_ As Integer
Dim COLUMNAAREA_ As Integer

'Para mover el frame
Dim OrigFX As Long
Dim OrigFY As Long

'*****************************************************************************************************
'* Nombre           : Filtrar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA UN FILTRO SOBRE EL RECORDSET RstFrm
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Filtrar()
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String

    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":        xCampos(0, 2) = "C":         xCampos(0, 3) = "4200"
    xCampos(1, 0) = "Codigo":       xCampos(1, 1) = "codrec":             xCampos(1, 2) = "C":         xCampos(1, 3) = "1500"

    TabOne1.CurrTab = 0
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg1
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

Private Sub cmd_Click(Index As Integer)
    Dim xRs As New ADODB.Recordset
    Dim xCampos() As String
    Dim A As Integer
    Dim SELECCIONADO As Double
    Dim DESCRIPCION As String
    Dim IDREC As Double
    Dim IDTAR As Double
    Dim IDUNIMED As Double
    Dim FILA As Integer
    Dim Rpta As Integer

    Dim SELECCIONADO_AUX As Double
    Dim DESCRIPCION_AUX As String
    Dim IDREC_AUX As Double
    Dim IDTAR_AUX As Double
    Dim IDUNIMED_AUX As Double
    
    Dim nSQLId As String
    Dim xform As New eps_librerias.FormSeleccion

    If QueHace = 3 Then Exit Sub

    Select Case Index
        Case 0 ' Elegir Receta
            ReDim xCampos(3, 4) As String
            Dim nTitulo As String
            Dim xRsAux As New ADODB.Recordset
            
            Set xRs = Nothing
            
            nTitulo = "Recetas"
            
            ' generar la lista de recetas para no considerar en la lista
            cSQL = "SELECT pro_personal.id, pro_personal.idrec, pro_receta.codrec, pro_receta.descripcion, pro_personal.activo " _
                + vbCr + "FROM pro_personal LEFT JOIN pro_receta ON pro_personal.idrec = pro_receta.id " _
                + vbCr + "GROUP BY pro_personal.id, pro_personal.idrec, pro_receta.codrec, pro_receta.descripcion, pro_personal.activo;"
            
            RST_Busq xRsAux, cSQL, xCon
            ' Se verifica que no se agregue una receta ya existente
            nSQLId = GENERAR_SQL_ID_RST(xRsAux, "idrec", " AND pro_receta.id", "NOT IN", True)

            cSQL = "SELECT pro_receta.id, pro_receta.codrec, alm_inventario.descripcion, mae_familia.descripcion AS desfam, IIf([prirec]=1,'PRINCIPAL','AUXILIAR') AS prioridad " _
                + vbCr + "FROM pro_receta LEFT JOIN (alm_inventario LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id) ON pro_receta.iditem = alm_inventario.id " _
                + vbCr + "Where (((alm_inventario.tippro) = 3)) " & nSQLId _
                + vbCr + "GROUP BY pro_receta.id, pro_receta.codrec, alm_inventario.descripcion, mae_familia.descripcion, IIf([prirec]=1,'PRINCIPAL','AUXILIAR');"
            
            RST_Busq xRs, cSQL, xCon
            
            'descripcion                        'campo                           'tamaño                    'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":    xCampos(0, 3) = "C"
            xCampos(1, 0) = "Cod. Rec":         xCampos(1, 1) = "codrec":        xCampos(1, 2) = "1000":    xCampos(1, 3) = "C"
            xCampos(2, 0) = "Prioridad":        xCampos(2, 1) = "prioridad":     xCampos(2, 2) = "2000":    xCampos(2, 3) = "C"

            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), "Buscando " & nTitulo, "descripcion", "descripcion", Principio

            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            If NulosC(TxtCodRec.Text) <> "" And RstValores.RecordCount <> 0 Then
                If MsgBox("Esta seguro que desea cambiar de Receta, se eliminara todo el pesonal relcionado a la Receta anterior?" _
                                    , vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
            End If
            
            limpiarRST RstValores, True
            fg(0).Rows = fg(0).FixedRows
            fg(1).Rows = fg(1).FixedRows

            TxtCodRec.Text = NulosC(xRs("codrec")) ' CODIGO DE LA RECETA
            lblIdRec.Caption = NulosN(xRs("id")) ' ID DE LA RECETA
            LblDesRec = NulosC(xRs("descripcion")) ' DESCRIPCION DE LA RECETA
            Agregando = True
            mostrarTareas (NulosN(xRs("id")))
            '--determinar cantidad de personal por tarea
            hallarCantPers
            Agregando = False
            
        Case 1 ' Agregar Personal
            ReDim xCampos(4, 4) As String
            
            xCampos(0, 0) = "Num. Doc.":            xCampos(0, 1) = "numdoc":       xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Apellidos y Nombres":  xCampos(1, 1) = "nombre":       xCampos(1, 2) = "4000":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
            xCampos(2, 0) = "Fch. Ing.":            xCampos(2, 1) = "fching":       xCampos(2, 2) = "1000":     xCampos(2, 3) = "C":    xCampos(2, 4) = "C"
            xCampos(3, 0) = "Area":                 xCampos(3, 1) = "area":         xCampos(3, 2) = "2000":     xCampos(3, 3) = "C":    xCampos(3, 4) = "C"
            
            If fg(0).Rows <= fg(0).FixedRows Then
                MsgBox "Agregue una Receta adecuada para procesar el Personal", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            
            ' generar la lista de personal para no considerar en la lista
            RstValores.Filter = "idtar = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNAIDTAR_)) & ""
            nSQLId = GENERAR_SQL_ID_RST(RstValores, "idpers", " AND pla_empleados.id", "NOT IN", True)
            
            ' generar la consulta
            cSQL = "SELECT pla_empleados.id AS idemp, pla_empleados.nombre, pla_empleados.numdoc, pla_empleados.fching, pla_empleados.fchcese, mae_area.descripcion AS area " _
                + vbCr + "FROM ((pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN mae_area ON pla_empleados.idarea = mae_area.id " _
                + vbCr + "WHERE  (((pla_empleados.fchcese) Is Null) AND ((pro_empdet.idfun)=6)) " & nSQLId _
                + vbCr + "ORDER BY pla_empleados.nombre;"
                
            nTitulo = "Buscando Personal"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
                            
            If xRs.State = 0 Then Exit Sub
            
            ' agregando los datos al rst temporal
            RstValores.AddNew
            RstValores("idtar") = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNAIDTAR_))
            RstValores("idpers") = NulosN(xRs("idemp"))
            RstValores("nombre") = NulosC(xRs("nombre"))
            RstValores("numdoc") = NulosC(xRs("numdoc"))
            RstValores("fching") = xRs("fching")
            RstValores("fchcese") = xRs("fchcese")
            RstValores("area") = NulosC(xRs("area"))
            RstValores("activo") = -1
            RstValores("estado") = -1
            RstValores.Update
            ' Se agrega el numero de orden
            agregarOrden True, NulosN(xRs("idemp"))
            
            pCargarDatosValores NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNAIDTAR_))
            
            Agregando = False
            Set xRs = Nothing

        Case 2 ' Listar Personal
            ReDim xCampos(4, 4) As String
            
            xCampos(0, 0) = "Num. Doc":             xCampos(0, 1) = "numdoc":      xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Apellidos y Nombres":  xCampos(1, 1) = "nombre":      xCampos(1, 2) = "4500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
            xCampos(2, 0) = "Fch. Ing.":            xCampos(2, 1) = "fching":      xCampos(2, 2) = "1000":     xCampos(2, 3) = "D":    xCampos(2, 4) = "D"
            xCampos(3, 0) = "Area":                 xCampos(3, 1) = "area":        xCampos(3, 2) = "2000":     xCampos(3, 3) = "C":    xCampos(3, 4) = "C"
            
            If fg(0).Rows <= fg(0).FixedRows Then
                MsgBox "Agregue una Receta adecuada para procesar el Personal", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            
            ' generar la lista de personal para no considerar en la lista
            RstValores.Filter = "idtar = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNAIDTAR_)) & ""
            nSQLId = GENERAR_SQL_ID_RST(RstValores, "idpers", " AND pla_empleados.id", "NOT IN", True)
            
            ' generar la consulta
            cSQL = "SELECT 0 AS xsel, pla_empleados.id AS idemp, pla_empleados.nombre, pla_empleados.numdoc, pla_empleados.fching, pla_empleados.fchcese, mae_area.descripcion AS area " _
                + vbCr + "FROM ((pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN mae_area ON pla_empleados.idarea = mae_area.id " _
                + vbCr + "WHERE  (((pla_empleados.fchcese) Is Null) AND ((pro_empdet.idfun)=6)) " & nSQLId _
                + vbCr + "ORDER BY pla_empleados.nombre; "
                
            nTitulo = "Buscando Personal"
        
            xform.SQLCad = cSQL
                
            xform.titulo = "Buscando Personal"
            Set xform.Coneccion = xCon
            Set xRs = Nothing
            Set xRs = xform.seleccionar(xCampos)
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            For A = 1 To xRs.RecordCount
                ' agregando los datos al rst temporal
                RstValores.AddNew
                RstValores("idtar") = NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNAIDTAR_))
                RstValores("idpers") = NulosN(xRs("idemp"))
                RstValores("nombre") = NulosC(xRs("nombre"))
                RstValores("numdoc") = NulosC(xRs("numdoc"))
                RstValores("fching") = xRs("fching")
                RstValores("fchcese") = xRs("fchcese")
                RstValores("area") = NulosC(xRs("area"))
                RstValores("activo") = -1
                RstValores("estado") = -1
                RstValores.Update
                
                ' Se agrega el numero de orden
                agregarOrden True, NulosN(xRs("idemp"))
                xRs.MoveNext
                If xRs.EOF = True Then Exit For
            Next A
            
            pCargarDatosValores NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNAIDTAR_))
            
            Set xform = Nothing
            Set xRs = Nothing
            
        Case 3 ' Eliminar Personal
            If fg(1).Row < 1 Then
                MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                fg(0).SetFocus
                Exit Sub
            End If
            
            If fg(1).Rows = 1 Then
                MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                fg(1).SetFocus
                Exit Sub
            End If
            
            If Not ELIMINARTODO_ Then
                If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
            End If
            
            RstValores.Filter = "idtar = " & fg(0).TextMatrix(fg(0).Row, COLUMNAIDTAR_)
            If RstValores.RecordCount = 0 Then Exit Sub
            
            RstValores.MoveFirst
            Do While Not RstValores.EOF
                If RstValores.RecordCount = 0 Then Exit Do
                If NulosN(RstValores("idpers")) = NulosN(fg(1).TextMatrix(fg(1).Row, COLUMNAIDPERS_)) Then
                    RstValores.Delete
                    Exit Do
                End If
                RstValores.MoveNext
            Loop
            
            pCargarDatosValores NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNAIDTAR_))
            
        Case 4 ' Eliminar Lista Personal
            If fg(1).Row < 1 Then
                MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                fg(0).SetFocus
                Exit Sub
            End If
            
            If fg(1).Rows = 1 Then
                MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                fg(1).SetFocus
                Exit Sub
            End If
            
            If MsgBox("Esta seguro desea eliminar toda la lista?", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
            
            For A = 1 To fg(1).Rows - 1
                ELIMINARTODO_ = True
                fg(1).Select 1, 1, 1, fg(1).Cols - 1
                cmd_Click 3
            Next A
            pCargarDatosValores NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNAIDTAR_))
            
            ELIMINARTODO_ = False
            
    End Select
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstFrm
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA EN FORMA ASCENDENTE O DECENDENTE LAS COLUMNAS DEL CONTROL Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstFrm("id")), xCon
    End If
End Sub

Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Agregando Then Exit Sub
    If Index = 1 Then
        Select Case Col
            Case COLUMNAORDEN_
                Dim IDPERSONAL_ As Double
                
                IDPERSONAL_ = NulosN(fg(1).TextMatrix(fg(1).Row, COLUMNAIDPERS_))
                RstValores.Filter = "idtar = " & NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNAIDTAR_)) & _
                                        " And idpers = " & NulosN(fg(1).TextMatrix(fg(1).Row, COLUMNAIDPERS_))
                
                If RstValores.State = 0 Then Exit Sub
                If RstValores.RecordCount = 0 Then Exit Sub
        
                RstValores("orden") = NulosN(fg(1).TextMatrix(fg(1).Row, COLUMNAORDEN_))
                RstValores.Update
                
                pCargarDatosValores fg(0).TextMatrix(fg(0).Row, COLUMNAIDTAR_)
                buscarPersonal IDPERSONAL_
        End Select
    End If
End Sub

Private Sub Fg_EnterCell(Index As Integer)
    Select Case Index
        Case 0
            fg(Index).Editable = flexEDNone
            fg(Index).SelectionMode = flexSelectionByRow
            
        Case 1
            fg(Index).SelectionMode = flexSelectionFree
            If QueHace = 3 Then
                fg(Index).Editable = flexEDNone
                fg(Index).AutoSearch = flexSearchFromTop
                Exit Sub
            End If
            fg(Index).Editable = flexEDKbdMouse
            fg(Index).AutoSearch = flexSearchNone
            
    End Select
End Sub

Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If Index = 1 Then
        Select Case Col
            Case COLUMNAORDEN_
                If validar_numero(KeyAscii) = False Then KeyAscii = 0
                Exit Sub
        End Select
    End If
    KeyAscii = 0
End Sub

Private Sub Fg_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo error
    If Index = 0 Then Exit Sub
    If KeyCode = 114 Or KeyCode = vbKeyInsert Then cmd_Click 1      'F3 = Agregar Item
    If KeyCode = 115 Or KeyCode = vbKeyDelete Then cmd_Click 3      'F4 = Eliminar Item
    Exit Sub

error:
    SHOW_ERROR Me.Name, "Fg_KeyUp (" & Index & ")"
End Sub

Private Sub fg_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Agregando = True Then Exit Sub
    If QueHace = 3 Then Exit Sub
    If Index = 0 Then Exit Sub

    Dim FILASELECCIONADA_ As Integer
    Dim COLUMNASELECCIONADA_ As Integer
    Dim FILATOPE_ As Integer
    
    If Index = 1 Then
        RstValores.Filter = "idtar = " & fg(0).TextMatrix(fg(0).Row, COLUMNAIDTAR_) & _
                                " And idpers = " & fg(1).TextMatrix(fg(1).Row, COLUMNAIDPERS_)
        FILASELECCIONADA_ = fg(1).Row
        COLUMNASELECCIONADA_ = fg(1).Col
        FILATOPE_ = fg(1).TopRow
        
        If RstValores.State = 0 Then Exit Sub
        If RstValores.RecordCount = 0 Then Exit Sub

        Select Case fg(1).Col
            Case COLUMNAACTIVO_ ' ACTIVO
                RstValores("activo") = NulosN(fg(1).TextMatrix(fg(1).Row, COLUMNAACTIVO_))
        End Select
        RstValores.Update
        pCargarDatosValores fg(0).TextMatrix(fg(0).Row, COLUMNAIDTAR_)
        fg(1).Select FILASELECCIONADA_, COLUMNASELECCIONADA_
        fg(1).TopRow = FILATOPE_
    End If
End Sub

Private Sub fg_RowColChange(Index As Integer)
    If Index = 1 Then Exit Sub

    If fg(0).Row < 1 Then
        fg(1).Rows = 1
        Exit Sub
    End If

    If RstValores.State = 0 Then Exit Sub

    ' Mostramos el Personal de la Tarea seleccionada
    pCargarDatosValores NulosN(fg(0).TextMatrix(fg(0).Row, COLUMNAIDTAR_))
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    On Error GoTo error
    If SeEjecuto = False Then

        SeEjecuto = True

        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu

        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        cSQL = "SELECT pro_personal.id, pro_personal.idrec, pro_receta.codrec, pro_receta.descripcion, pro_personal.activo " _
            + vbCr + "FROM pro_personal LEFT JOIN pro_receta ON pro_personal.idrec = pro_receta.id " _
            + vbCr + "GROUP BY pro_personal.id, pro_personal.idrec, pro_receta.codrec, pro_receta.descripcion, pro_personal.activo;"

        RST_Busq RstFrm, cSQL, xCon

        Set Dg1.DataSource = RstFrm
    End If
    Exit Sub
error:
    SHOW_ERROR Me.Name, "Form_Activate"
End Sub

Private Sub iniciarCampos()
    TabOne1.CurrTab = 0
    fg(0).SelectionMode = flexSelectionByRow
    
    fg(1).AllowUserResizing = flexResizeColumns
    fg(1).AutoSearch = flexSearchFromTop
    fg(1).ExplorerBar = flexExSortShow
    
    ELIMINARTODO_ = False
    ' SE INICIALIZAN LOS VALORES DE LAS COLUMNAS
    COLUMNADETTAREA_ = 1
    COLUMNATOTPERS_ = 2
    COLUMNAPERSSEL_ = 3
    COLUMNAPERSACT_ = 4
    COLUMNAIDTAR_ = 5
    
    COLUMNAORDEN_ = 1
    COLUMNAACTIVO_ = 2
    COLUMNADETPERS_ = 3
    COLUMNANUMDOC_ = 4
    COLUMNAFCHINGR_ = 5
    COLUMNAFCHCESE_ = 6
    COLUMNAAREA_ = 7
    COLUMNAIDPERS_ = 8
    
    fg(0).FrozenCols = COLUMNADETTAREA_
    fg(1).FrozenCols = COLUMNADETPERS_
    
    ' ocultando las columnas de codigos
    OCULTAR_COL fg(0), COLUMNAIDTAR_, COLUMNAIDTAR_
    OCULTAR_COL fg(1), COLUMNAIDPERS_, COLUMNAIDPERS_
    OCULTAR_COL fg(1), COLUMNAACTIVO_, COLUMNAACTIVO_
End Sub

'*****************************************************************************************************
'* Nombre           : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Nuevo()
    Bloquea False
    Blanquea
    ActivaTool
    QueHace = 1
    xHorIni = Time
    Label5.Caption = "Agregando Personal"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False

    If RstValores.State = 0 Then pCargarDatosRstTemp 0
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS BOTONES DE LA BARRA DE HERRAMIENTAS
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
'* Descripcion      : PREPARA LOS CONTROLES TEXTBOX PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    TxtCodRec.Text = ""
    LblDesRec.Caption = ""
    lblIdRec.Caption = ""
    fg(0).Rows = 1
    fg(1).Rows = 1
End Sub

'*****************************************************************************************************
'* Nombre           : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES TEXTBOX Y COMMAND
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Bloquea(band As Boolean)
    TxtCodRec.Locked = band
    habilitar Cmd, Not band
    If Not band Then
        fg(0).Editable = flexEDKbdMouse
    Else
        fg(0).Editable = flexEDNone
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    QueHace = 3
    SeEjecuto = False
    Agregando = False
    iniciarCampos
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO EN LA PESTAÑA DETALLE DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    Blanquea
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then Exit Sub

    MuestraDetalle
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL PROCESO DE AGREGAR O MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    Bloquea True
    ActivaTool
    Label5.Caption = "Detalle del Personal"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Dg1.SetFocus
    limpiarRST RstValores
End Sub

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub

    If Me.Width <= 12000 Then Me.Width = 12000
    If Me.Height <= 8100 Then Me.Height = 8100

    ' Se dimensiona el Contenido
    TabOne1.Width = Me.Width - 100
    TabOne1.Height = Me.Height - 780
    
    Label4(0).Width = Me.Width - 100
    Dg1.Width = TabOne1.Width - 150
    Dg1.Height = TabOne1.Height - 960
        
    ' Se dimensiona el Detalle
    ' DETALLE DE RECETA
    Label5.Width = Me.Width - 100
    
    FrmReceta.Top = TabOne1.Top + 100
    FrmReceta.Width = TabOne1.Width - 225
    LblDesRec.Width = FrmReceta.Width - 2910
    
    'LblDetalle.Width = FrmReceta.Width - 2300
    
    ' DESCRIPCION DE TAREA
    frmTarea.Width = TabOne1.Width - 2025
    
    fg(0).Width = frmTarea.Width - 225
    
    ' DETALLE PERSONAL
    frmPersonal.Width = TabOne1.Width - 200
    frmPersonal.Height = TabOne1.Height - 3450
    
    fg(1).Width = frmPersonal.Width - 2010
    fg(1).Height = frmPersonal.Height - 345
    FrmBotones.Left = frmPersonal.Width - 1800
    FrmBotones.Height = frmPersonal.Height - 240
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    If QueHace <> 3 Then
'        MsgBox "No puede salir del formulario mientras este ingresando o modificando un Costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Cancel = True
'        Exit Sub
'    Else
'        Set RstFrm = Nothing
'        SeEjecuto = False
'    End If
'End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If RstFrm.RecordCount = 0 And QueHace <> 1 Then
            'xTitulo = "Mostrar Detalle"
            MsgBox "No hay informacion para mostrar, haga clic en Nuevo para agregar información", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Cancel = 1
            Exit Sub
        End If
        If QueHace = 3 Then MuestraSegundoTab
    Else
        limpiarRST RstValores, True
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo

    If Button.Index = 2 Then Modificar

    If Button.Index = 3 Then Eliminar

    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstFrm.Requery
            Dg1.Refresh

            RstFrm.MoveFirst
            RstFrm.Find "id = " & mIdRegistro & ""
            If RstFrm.EOF = True Then
                RstFrm.MoveFirst
            End If
        End If
    End If

    If Button.Index = 6 Then Cancelar

    If Button.Index = 8 Then Filtrar

    If Button.Index = 9 Then
        If RstFrm.State = 0 Then Exit Sub
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstFrm.Filter = adFilterNone
        RstFrm.Requery
        Dg1.Refresh
    End If

    If Button.Index = 10 Then Buscar

    If Button.Index = 12 Then pExportar

    If Button.Index = 18 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

Private Sub agregarOrden(AGREGAR_ As Boolean, IDPERS_ As Double, Optional NUMEROORDEN_ As Double)
    Dim A As Integer
    Dim NUMORDENAUX_ As Double
    
    If AGREGAR_ Then
        ' Se filtra el personal de la tarea seleccionada
        RstValores.Filter = adFilterNone
        RstValores.Filter = "idtar = " & fg(0).TextMatrix(fg(0).Row, COLUMNAIDTAR_)
        RstValores.Sort = "orden"
        RstValores.MoveLast
        NUMORDENAUX_ = NulosN(RstValores("orden"))
        
        RstValores.Filter = "idtar = " & fg(0).TextMatrix(fg(0).Row, COLUMNAIDTAR_) & " And idpers = " & IDPERS_
        RstValores("orden") = NUMORDENAUX_ + 1
    Else
        ' Se filtra el personal de la tarea seleccionada
        RstValores.Filter = adFilterNone
        RstValores.Filter = "idtar = " & fg(0).TextMatrix(fg(0).Row, COLUMNAIDTAR_)
        
        RstValores.Sort = "orden"
        NUMORDENAUX_ = NUMEROORDEN_ + 1
        RstValores.MoveFirst
        For A = 1 To RstValores.RecordCount
            If RstValores("idpers") = IDPERS_ Then
                GoTo SIGUIENTE
            End If
            RstValores("orden") = NUMORDENAUX_
            NUMORDENAUX_ = NUMORDENAUX_ + 1
SIGUIENTE:
            RstValores.MoveNext
            If RstValores.EOF Then Exit For
        Next A
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : FUNCION
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA pro_costodet, ESTA FUNCION DEVUELVE VERDADERO
'*                    CUANDO TIENE EXITO
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Function Grabar() As Boolean
    If fValidarDatos() = False Then
        MsgBox "No se pudo realizar la Operacion, ingrese correctamente los Datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Function
    End If

    If MsgBox("¿Seguro que desea " + IIf(QueHace = 1, "grabar", "Modificar") + " el Producto?", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo SALIR

    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xId As Double
    Dim xCol&, xFil&, xCorr&
    Dim A As Integer


On Error GoTo LaCague

    xCon.BeginTrans
    Me.MousePointer = vbHourglass

    If QueHace = 1 Then
        xId = HallaCodigoTabla("pro_personal", xCon, "id")
        RST_Busq RstCab, "SELECT top 1 * FROM pro_personal ", xCon
    Else
        xId = NulosN(RstFrm("id"))

        xCon.Execute "DELETE * FROM pro_personaldet WHERE idpropers = " & xId & ""
        xCon.Execute "DELETE * FROM pro_personal WHERE id = " & xId & ""

        RST_Busq RstCab, "SELECT top 1 * FROM pro_personal ", xCon
    End If

    mIdRegistro = xId

    RST_Busq RstDet, "SELECT top 1 * FROM pro_personaldet", xCon

    ' Recorrer cabeceras
    RstCab.AddNew
    RstCab("id") = xId
    RstCab("idrec") = NulosN(lblIdRec.Caption)
    RstCab("activo") = True
    RstCab.Update

    RstValores.Filter = adFilterNone
    If RstValores.RecordCount = 0 Then GoTo LaCague
    
    Dim IDDET_ As Double
    IDDET_ = HallaCodigoTabla("pro_personaldet", xCon, "id")
    
    RstValores.MoveFirst
    While Not RstValores.EOF
        RstDet.AddNew
        RstDet("id") = IDDET_
        RstDet("idpropers") = xId
        RstDet("idtar") = NulosN(RstValores("idtar"))
        RstDet("idpers") = NulosN(RstValores("idpers"))
        RstDet("orden") = NulosN(RstValores("orden"))
        RstDet("activo") = NulosN(RstValores("activo"))
        RstDet.Update
        RstValores.MoveNext
        IDDET_ = IDDET_ + 1
    Wend

    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId

    xCon.CommitTrans
    'xTitulo = "Grabar"
    MsgBox "El Producto se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    Grabar = True

SALIR:
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing
    Exit Function

LaCague:
    xCon.RollbackTrans
    Resume
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
    Grabar = False
End Function

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA LA EDICION DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    Bloquea False
    Blanquea
    ActivaTool
    QueHace = 2
    xHorIni = Time
    Label5.Caption = "Modificando Producto"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    limpiarRST RstValores
    MuestraSegundoTab
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA pro_personal
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    Dim RstTmp  As New ADODB.Recordset
    Dim nSQL As String
    Dim xId&

    TabOne1.CurrTab = 0
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay registros para eliminar", vbInformation, xTitulo
        Exit Sub
    End If

    xId = NulosN(RstFrm.Fields("id"))

    Set RstTmp = Nothing
    Rpta = MsgBox("Esta seguro de eliminar El Producto Seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM pro_personaldet WHERE idpropers =" & xId & " "
        xCon.Execute "DELETE * FROM pro_personal WHERE id =" & xId & " "
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & xId & " AND idform = " & IdMenuActivo

        MsgBox "El Personal relacionado al Producto se eliminó con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstFrm.Requery
        Dg1.Refresh
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA UNA BUSQUEDA EN EL RECORDSET RstFrm
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Buscar()
    On Error GoTo error
    Dim RstTmp As New ADODB.Recordset
    Dim xCampos(3, 4) As String

    xCampos(0, 0) = "Tipo":           xCampos(0, 1) = "Origen":       xCampos(0, 2) = "1200":        xCampos(0, 3) = "c"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "Referencia":   xCampos(1, 2) = "4500":        xCampos(1, 3) = "c"
    xCampos(2, 0) = "Cod.Rec":        xCampos(2, 1) = "codrec":       xCampos(2, 2) = "900":         xCampos(2, 3) = "c"

    TabOne1.CurrTab = 0

    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, RstFrm.Source, xCampos(), "Buscando Costo", "Referencia", "Referencia", Principio
    If RstTmp.State = 0 Then GoTo SALIR
    If RstTmp.EOF = True And RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR
    RstFrm.MoveFirst
    RstFrm.Find "id = " & RstTmp("id") & ""

SALIR:
    Set RstTmp = Nothing

error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub

'*****************************************************************************************************
'* Nombre           : fValidarDatos
'* Tipo             : FUNCCION
'* Descripcion      : VERIFICA QUE LOS DATOS INGRESADO SEAN LOS CORRECTOS, ESTA FUNCION DEVUELVE
'*                    VERDADERO
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Private Function fValidarDatos() As Boolean
    Dim A As Integer
    Dim valor As Boolean
    Dim SELECCIONADO As Double

    valor = True

    If NulosC(LblDesRec.Caption) = "" Then valor = False

    If RstValores.State = 0 Then valor = False
    RstValores.Filter = adFilterNone
    If RstValores.RecordCount = 0 Then valor = False
    
    fValidarDatos = valor
End Function

'*****************************************************************************************************
'* Nombre           : pExportar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA A MS EXCEL LOS DATOS DEL RECORDSET RSTTMP
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pExportar()
    TabOne1.CurrTab = 0

    Dim xCampos(3, 3) As String
    Dim oExport As New SGI2_funciones.formularios
    Dim RstTmp As New ADODB.Recordset
    Set RstTmp = RstFrm.Clone
    ' 0 Nombre a Mostrar;
    ' 1 nombre de Campo del Rst;
    ' 2 alineacion(0::derecha, 1::centro, 2::izquierda);
    ' 3 ancho de columna
    ' obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Código":       xCampos(0, 1) = "codigo":       xCampos(0, 2) = 0:  xCampos(0, 3) = "1200"
    xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = 0:  xCampos(1, 3) = "3500"
    xCampos(2, 0) = "Unidad":       xCampos(2, 1) = "abrev":        xCampos(2, 2) = 0:  xCampos(2, 3) = "750"
    xCampos(3, 0) = "Es Diverso":   xCampos(3, 1) = "diverso":      xCampos(3, 2) = 0:  xCampos(3, 3) = "800"

    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "Lista de Personal", "", "", "Listado de Personal", RstTmp, xCampos()

    Set oExport = Nothing
    Set RstTmp = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraDetalle
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO ACTUAL
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub MuestraDetalle()
    Dim RstTmp As New ADODB.Recordset

    On Error GoTo error

    TxtCodRec.Text = NulosC(RstFrm("codrec"))
    LblDesRec.Caption = NulosC(RstFrm("descripcion"))
    lblIdRec.Caption = NulosN(RstFrm("idrec"))
    
    ' Se llenan las Tareas de esa Receta
    mostrarTareas (NulosN(RstFrm("idrec")))
    ' Se llena el personal relacionado a las tareas de la Receta
    fg(1).Rows = 1
    pCargarDatosRstTemp NulosN(RstFrm("id"))
    pCargarDatosValores fg(0).TextMatrix(fg(0).Row, COLUMNAIDTAR_)
    
    Set RstTmp = Nothing
    Agregando = False
    Exit Sub
error:
    Set RstTmp = Nothing
    Agregando = False
    SHOW_ERROR Me.Name, "MuestraDetalle"
End Sub

Private Sub buscarPersonal(IDPERS_ As Double)
    Dim A As Integer
    Dim FILASEL_ As Integer
    
    For A = 1 To fg(1).Rows - 1
        If fg(1).TextMatrix(A, COLUMNAIDPERS_) = IDPERS_ Then
            FILASEL_ = A
            Exit For
        End If
    Next A
    
    fg(1).Row = FILASEL_
    fg(1).TopRow = FILASEL_
End Sub

'*****************************************************************************************************
'* Nombre           : mostrarTareas
'* Tipo             : SUB
'* Descripcion      : Muestra las Tareas de la Receta seleccionada
'* Parametros       : IDREC_: Identificador de la Receta
'* Devuelve         :
'*****************************************************************************************************
Private Sub mostrarTareas(IDREC_ As Double)
    Dim RstTmp As New ADODB.Recordset
    Dim A As Integer
    
    cSQL = "SELECT pro_tareas.id AS idtar, pro_tareas.descripcion AS destar " _
        + vbCr + "FROM pro_recetatar LEFT JOIN pro_tareas ON pro_recetatar.idtar = pro_tareas.id " _
        + vbCr + "Where (((pro_recetatar.idrec) = " & IDREC_ & ")) " _
        + vbCr + "GROUP BY pro_tareas.id, pro_tareas.descripcion " _
        + vbCr + "ORDER BY pro_tareas.descripcion;"
        
    RST_Busq RstTmp, cSQL, xCon
    
    fg(0).Rows = 1
    If RstTmp.State = 0 Then Exit Sub
    If RstTmp.RecordCount = 0 Then Exit Sub
    
    RstTmp.MoveFirst
    For A = 1 To RstTmp.RecordCount
        fg(0).Rows = fg(0).Rows + 1
        fg(0).TextMatrix(A, COLUMNADETTAREA_) = RstTmp("destar")
        fg(0).TextMatrix(A, COLUMNAIDTAR_) = RstTmp("idtar")
        
        RstTmp.MoveNext
    Next A
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarDatosRstTemp
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : Definir la estructura del recordset de los valores, ESTA FUNCION DEVUELVE UN
'*                    RECORDSET CON DATOS
'* Paranetros       : NOMBRE    |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    idCodigo  |  INTEGER    |  codigo del Costo
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarDatosRstTemp(idCodigo)
    Dim RstTmp As New ADODB.Recordset
    Set RstTmp = Nothing
    'campo estado=Indica si personal esta activo o cesado, -1=Activo; 0=Cesado
    ' definir la estructura de recordset
    cSQL = "SELECT pro_personaldet.idtar, pro_personaldet.idpers, pla_empleados.nombre, pla_empleados.numdoc, pro_personaldet.activo, IIf([fchcese] Is Null ,-1,0) AS estado, pro_personaldet.orden, pla_empleados.fching, pla_empleados.fchcese, mae_area.descripcion AS area " _
        + vbCr + "FROM (pro_personaldet LEFT JOIN pla_empleados ON pro_personaldet.idpers = pla_empleados.id) LEFT JOIN mae_area ON pla_empleados.idarea = mae_area.id " _
        + vbCr + "Where (((pro_personaldet.idpropers) = " & idCodigo & ")) " _
        + vbCr + "ORDER BY pro_personaldet.orden;"

    RST_Busq RstTmp, cSQL, xCon

    If RstValores.State = 0 Then DEFINIR_RST_TMP RstValores, RstTmp
    CARGAR_RST_TMP RstValores, RstTmp

    Set RstTmp = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : hallarCantPers
'* Tipo             : SUB
'* Descripcion      : Limpia el numero de personas activas, seleccionadas y vigentes de las tareas
'* Parametros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub hallarCantPers()
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    Dim TOTALPERS_ As Double
    Dim PERSREC_ As Double
    Dim PERSACT_ As Double
    
    If RstValores.State = 0 Then Exit Sub
    
    For A = 1 To fg(0).Rows - 1
        RstValores.Filter = "idtar = " & NulosN(fg(0).TextMatrix(A, COLUMNAIDTAR_))
        ' Se llena el total de Personas
        TOTALPERS_ = RstValores.RecordCount
        
        ' Se verifica el numero de personas recomendado segun linea
        cSQL = "SELECT pro_lineadet.numop " _
            + vbCr + "FROM pro_lineadet LEFT JOIN pro_linea ON pro_lineadet.idlineadet = pro_linea.id " _
            + vbCr + "WHERE (((pro_linea.activo)=-1) AND ((pro_linea.idrec)=" & NulosN(lblIdRec.Caption) & ") AND ((pro_lineadet.idtar)=" & NulosN(fg(0).TextMatrix(A, COLUMNAIDTAR_)) & "));"
        
        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 0 Then PERSREC_ = 0: GoTo SIGUIENTE
        If xRs.RecordCount = 0 Then PERSREC_ = 0: GoTo SIGUIENTE
        
        PERSREC_ = NulosN(xRs("numop"))
SIGUIENTE:
        ' se llena las personas activas
        RstValores.Filter = adFilterNone
        RstValores.Filter = "idtar = " & NulosN(fg(0).TextMatrix(A, COLUMNAIDTAR_)) & _
                                            " And activo = True" & _
                                            " And estado = -1"
        PERSACT_ = RstValores.RecordCount
        
        fg(0).TextMatrix(A, COLUMNATOTPERS_) = Format(TOTALPERS_, "00")
        fg(0).TextMatrix(A, COLUMNAPERSSEL_) = Format(PERSREC_, "00")
        fg(0).TextMatrix(A, COLUMNAPERSACT_) = Format(PERSACT_, "00")
    Next A
End Sub

'*****************************************************************************************************
'* Nombre           : limpiarRST
'* Tipo             : SUB
'* Descripcion      : Limpia el recordset segun parametros
'* Parametros       : Rst: Recordset a Limpiar
'*                    TODO: especifica si se tiene que limpiar todo el recordset o solo la parte filtrada
'* Devuelve         :
'*****************************************************************************************************
Private Sub limpiarRST(Rst As ADODB.Recordset, Optional TODO As Boolean = True)
    With Rst
        If .State <> 0 Then
            If TODO Then .Filter = adFilterNone
            If .RecordCount <> 0 Then
                .MoveFirst
                While Not .EOF
                    .Delete
                    .MoveNext
                Wend
            End If
        End If
    End With
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarDatosValores
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : Muestra el personal de la Tarea especificada
'* Parametros       : IDTAR_: Identificador de la Tarea a mostrar
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarDatosValores(IDTAR_ As Double)
    Agregando = True
    RstValores.Filter = adFilterNone
    RstValores.Filter = "idtar = " & IDTAR_ & ""
    fg(1).Rows = 1
    If RstValores.State = 0 Then GoTo SALIR
    If RstValores.RecordCount = 0 Then GoTo SALIR
    
    RstValores.Sort = "orden"
    RstValores.MoveFirst
    With fg(1)
        Do While Not RstValores.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, COLUMNAORDEN_) = Format(NulosN(RstValores("orden")), "00")
            .TextMatrix(.Rows - 1, COLUMNADETPERS_) = NulosC(RstValores("nombre"))
            .TextMatrix(.Rows - 1, COLUMNANUMDOC_) = NulosC(RstValores("numdoc"))
            .TextMatrix(.Rows - 1, COLUMNAFCHINGR_) = NulosC(Format(RstValores("fching"), "dd/mm/yyyy"))
            .TextMatrix(.Rows - 1, COLUMNAFCHCESE_) = NulosC(Format(RstValores("fchcese"), "dd/mm/yyyy"))
            .TextMatrix(.Rows - 1, COLUMNAAREA_) = NulosC(RstValores("area"))
            If RstValores("activo") Then .TextMatrix(.Rows - 1, COLUMNAACTIVO_) = -1 Else .TextMatrix(.Rows - 1, COLUMNAACTIVO_) = 0
            .TextMatrix(.Rows - 1, COLUMNAIDPERS_) = NulosN(RstValores("idpers"))
            
            .Select .Rows - 1, 1, .Rows - 1, .Cols - 1
            'campo estado=Indica si personal esta activo o cesado, -1=Activo; 0=Cesado
            If NulosN(RstValores("estado") = 0) Then
                .FillStyle = flexFillRepeat
                .CellBackColor = &HB9B9FF
            End If

            RstValores.MoveNext
        Loop
    .Select 1, 1
    End With
SALIR:
    hallarCantPers
    Agregando = False
End Sub
