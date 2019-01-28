VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManConfigVal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Configuracion de Valorizacion"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10905
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
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConfigVal.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConfigVal.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConfigVal.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConfigVal.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConfigVal.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConfigVal.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConfigVal.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConfigVal.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConfigVal.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConfigVal.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConfigVal.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConfigVal.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManConfigVal.frx":277E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6600
      Left            =   0
      TabIndex        =   0
      Top             =   375
      Width           =   10905
      _cx             =   19235
      _cy             =   11642
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
         Caption         =   "Frame1"
         Height          =   6180
         Left            =   11550
         TabIndex        =   4
         Top             =   375
         Width           =   10815
         Begin VB.Frame FrmReceta 
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
            Height          =   5775
            Left            =   60
            TabIndex        =   7
            Top             =   360
            Width           =   10680
            Begin VB.Frame Frame4 
               Caption         =   "[ Costos Indirectos ]"
               Height          =   1845
               Left            =   720
               TabIndex        =   22
               Top             =   3090
               Width           =   9195
               Begin VB.CheckBox ckprocmod 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Procesa MOD"
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   180
                  TabIndex        =   36
                  Top             =   360
                  Width           =   1425
               End
               Begin VB.CheckBox ckproccif 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Procesa CIF"
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   2670
                  TabIndex        =   35
                  Top             =   330
                  Width           =   1425
               End
               Begin VB.CommandButton cmd 
                  Enabled         =   0   'False
                  Height          =   240
                  Index           =   0
                  Left            =   2355
                  Picture         =   "FrmManConfigVal.frx":2B10
                  Style           =   1  'Graphical
                  TabIndex        =   25
                  Top             =   780
                  Width           =   240
               End
               Begin VB.CommandButton cmd 
                  Enabled         =   0   'False
                  Height          =   240
                  Index           =   1
                  Left            =   2355
                  Picture         =   "FrmManConfigVal.frx":2C42
                  Style           =   1  'Graphical
                  TabIndex        =   24
                  Top             =   1110
                  Width           =   240
               End
               Begin VB.CommandButton cmd 
                  Enabled         =   0   'False
                  Height          =   240
                  Index           =   2
                  Left            =   2355
                  Picture         =   "FrmManConfigVal.frx":2D74
                  Style           =   1  'Graphical
                  TabIndex        =   23
                  Top             =   1440
                  Width           =   240
               End
               Begin VB.TextBox txttipdist 
                  Height          =   300
                  Left            =   1410
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   26
                  Text            =   "txttipdist"
                  Top             =   750
                  Width           =   1215
               End
               Begin VB.TextBox txttipitemdist 
                  Height          =   300
                  Left            =   1410
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   27
                  Text            =   "txttipitemdist"
                  Top             =   1080
                  Width           =   1215
               End
               Begin VB.TextBox txtfacdist 
                  Height          =   300
                  Left            =   1410
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   28
                  Text            =   "txtfacdist"
                  Top             =   1410
                  Width           =   1215
               End
               Begin VB.Label lbl_cb_capt 
                  AutoSize        =   -1  'True
                  Caption         =   "Tip. Distrib."
                  Height          =   195
                  Index           =   1
                  Left            =   180
                  TabIndex        =   34
                  Top             =   780
                  Width           =   795
               End
               Begin VB.Label lbl_cb_capt 
                  AutoSize        =   -1  'True
                  Caption         =   "Tip. Item Distrib."
                  Height          =   195
                  Index           =   0
                  Left            =   180
                  TabIndex        =   33
                  Top             =   1140
                  Width           =   1140
               End
               Begin VB.Label lbl_cb_capt 
                  AutoSize        =   -1  'True
                  Caption         =   "Factor Distrib."
                  Height          =   195
                  Index           =   2
                  Left            =   180
                  TabIndex        =   32
                  Top             =   1470
                  Width           =   975
               End
               Begin VB.Label lbltipdist 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbltipdist"
                  ForeColor       =   &H00800000&
                  Height          =   300
                  Left            =   2670
                  TabIndex        =   31
                  Top             =   750
                  Width           =   6300
               End
               Begin VB.Label lbltipitemdist 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbltipitemdist"
                  ForeColor       =   &H00800000&
                  Height          =   300
                  Left            =   2670
                  TabIndex        =   30
                  Top             =   1080
                  Width           =   6300
               End
               Begin VB.Label lblfacdist 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lblfacdist"
                  ForeColor       =   &H00800000&
                  Height          =   300
                  Left            =   2670
                  TabIndex        =   29
                  Top             =   1410
                  Width           =   6300
               End
            End
            Begin VB.Frame Frame3 
               Caption         =   "[ Costo MP ]"
               Height          =   1365
               Left            =   720
               TabIndex        =   16
               Top             =   1620
               Width           =   9195
               Begin VB.CheckBox ckprocmp 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Procesa MP"
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   180
                  TabIndex        =   21
                  Top             =   390
                  Width           =   1425
               End
               Begin VB.CommandButton cmd 
                  Enabled         =   0   'False
                  Height          =   240
                  Index           =   3
                  Left            =   2355
                  Picture         =   "FrmManConfigVal.frx":2EA6
                  Style           =   1  'Graphical
                  TabIndex        =   17
                  Top             =   810
                  Width           =   240
               End
               Begin VB.TextBox txtmetval 
                  Height          =   300
                  Left            =   1410
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   18
                  Text            =   "txttipdist"
                  Top             =   780
                  Width           =   1215
               End
               Begin VB.Label lbl_cb_capt 
                  AutoSize        =   -1  'True
                  Caption         =   "Met. Val."
                  Height          =   195
                  Index           =   3
                  Left            =   180
                  TabIndex        =   20
                  Top             =   810
                  Width           =   630
               End
               Begin VB.Label lblmetval 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "lbltipdist"
                  ForeColor       =   &H00800000&
                  Height          =   300
                  Left            =   2670
                  TabIndex        =   19
                  Top             =   780
                  Width           =   6300
               End
            End
            Begin VB.CheckBox ActivoCheck 
               Caption         =   "Activo"
               Enabled         =   0   'False
               Height          =   255
               Left            =   9000
               TabIndex        =   12
               Top             =   990
               Width           =   855
            End
            Begin VB.TextBox DescripcionText 
               Height          =   300
               Left            =   1965
               Locked          =   -1  'True
               TabIndex        =   9
               Text            =   "DescripcionText"
               Top             =   1275
               Width           =   7935
            End
            Begin VB.TextBox AbrevText 
               Height          =   300
               Left            =   1965
               Locked          =   -1  'True
               TabIndex        =   8
               Text            =   "AbrevText"
               Top             =   960
               Width           =   2835
            End
            Begin VB.Label lblidmetval 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "lblidmetval"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   8640
               TabIndex        =   37
               Top             =   270
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblidfacdist 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "lblidfacdist"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   7770
               TabIndex        =   15
               Top             =   270
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblidtipitemdist 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "lblidtipitemdist"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   6660
               TabIndex        =   14
               Top             =   270
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.Label lblidtipdist 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "lblidtipdist"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   5850
               TabIndex        =   13
               Top             =   270
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Abrev."
               Height          =   195
               Index           =   0
               Left            =   720
               TabIndex        =   11
               Top             =   990
               Width           =   465
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Descripcion"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   10
               Top             =   1320
               Width           =   840
            End
         End
         Begin VB.Label TituloLabel 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Configuracion de Valorizacion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   105
            TabIndex        =   5
            Top             =   120
            Width           =   10665
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6180
         Left            =   45
         TabIndex        =   1
         Top             =   375
         Width           =   10815
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   5550
            Left            =   30
            TabIndex        =   2
            Top             =   585
            Width           =   10725
            _ExtentX        =   18918
            _ExtentY        =   9790
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
            Columns(1).Caption=   "Codigo"
            Columns(1).DataField=   "abrev"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descripcion"
            Columns(2).DataField=   "descripcion"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Met. Val."
            Columns(3).DataField=   "desmetval"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fact. Distrib."
            Columns(4).DataField=   "desfacdist"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Tip. Distrib."
            Columns(5).DataField=   "destipdist"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Tip. Item Distrib."
            Columns(6).DataField=   "destipitemdist"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Activo"
            Columns(7).DataField=   "activo"
            Columns(7).NumberFormat=   "Yes/No"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1429"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1349"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=4815"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=4736"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=2619"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2540"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=2461"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2381"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=2487"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2408"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=2461"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=2381"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=1429"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=1349"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=70,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
            _StyleDefs(68)  =   "Named:id=33:Normal"
            _StyleDefs(69)  =   ":id=33,.parent=0"
            _StyleDefs(70)  =   "Named:id=34:Heading"
            _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(72)  =   ":id=34,.wraptext=-1"
            _StyleDefs(73)  =   "Named:id=35:Footing"
            _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(75)  =   "Named:id=36:Selected"
            _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=37:Caption"
            _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(79)  =   "Named:id=38:HighlightRow"
            _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(81)  =   "Named:id=39:EvenRow"
            _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(83)  =   "Named:id=40:OddRow"
            _StyleDefs(84)  =   ":id=40,.parent=33"
            _StyleDefs(85)  =   "Named:id=41:RecordSelector"
            _StyleDefs(86)  =   ":id=41,.parent=34"
            _StyleDefs(87)  =   "Named:id=42:FilterBar"
            _StyleDefs(88)  =   ":id=42,.parent=33"
         End
         Begin VB.Label TituloLabel 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Configuracion de Valorizacion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   3
            Top             =   120
            Width           =   10590
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10905
      _ExtentX        =   19235
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
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Item"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar Item"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar un Item"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Retirar Item"
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
   Begin VB.Menu Menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "Eliminar             "
      End
   End
End
Attribute VB_Name = "FrmManConfigVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMANALAMACEN.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : AQUI SE CREAN, MODIFICAN Y ELIMINAN LOS ITEMS Y SE LES ASIGNA LA CUENTA CONTABLE.
'*                  : Y EL CENTRO DE COSTO.
'* DISEÑADO POR     : Jose Chacon
'* ULTIMA REVISION  : 24/04/2012
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim RstDesp As New ADODB.Recordset
Dim xHorIni As Date
Dim fOrdenLista As Boolean
Dim IdMenuActivo As Integer
Dim mIdRegistro&
Dim cSQL As String
Dim Agregando As Boolean
Dim F As New SistemaLogica.Funciones

Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Bloquea()
    AbrevText.Locked = Not AbrevText.Locked
    DescripcionText.Locked = Not DescripcionText.Locked
    ActivoCheck.Enabled = Not ActivoCheck.Enabled
    txtfacdist.Locked = Not txtfacdist.Locked
    txttipdist.Locked = Not txttipdist.Locked
    txttipitemdist.Locked = Not txttipitemdist.Locked
    habilitar cmd, Not cmd(0).Enabled
    ckprocmp.Enabled = Not ckprocmp.Enabled
    ckprocmod.Enabled = Not ckprocmod.Enabled
    ckproccif.Enabled = Not ckproccif.Enabled
End Sub

Private Sub iniciarCampos()
    Dim RstFacDist As New ADODB.Recordset
    
    TabOne1.CurrTab = 0
End Sub

Private Sub pCargarGrid()
    cSQL = "SELECT mae_configval.id, mae_configval.descripcion, mae_configval.abrev, mae_configval.activo, mae_configval.idmetval, mae_metodoval.codigo AS codmetval, mae_metodoval.descripcion AS desmetval, mae_configval.idfacdist, mae_metodovalfacdist.codigo AS codfacdist, mae_metodovalfacdist.descripcion AS desfacdist, mae_metodovalfacdist.columna AS colfacdist, mae_configval.idtipdist, mae_metodovaltipdist.codigo AS codtipdist, mae_metodovaltipdist.descripcion AS destipdist, mae_configval.idtipitemdist, mae_metodovaltipitemdist.codigo AS codtipitemdist, mae_metodovaltipitemdist.descripcion AS destipitemdist, mae_configval.procesamp, mae_configval.procesamod, mae_configval.procesacif " _
        + vbCr + "FROM (((mae_configval LEFT JOIN mae_metodovalfacdist ON mae_configval.idfacdist = mae_metodovalfacdist.id) LEFT JOIN mae_metodovaltipdist ON mae_configval.idtipdist = mae_metodovaltipdist.id) LEFT JOIN mae_metodovaltipitemdist ON mae_configval.idtipitemdist = mae_metodovaltipitemdist.id) LEFT JOIN mae_metodoval ON mae_configval.idmetval = mae_metodoval.id " _
        + vbCr + "ORDER BY mae_configval.descripcion"
    
    RST_Busq RstDesp, cSQL, xCon

    Set Dg1.DataSource = RstDesp
End Sub

Sub Blanquea()
    AbrevText.Text = ""
    DescripcionText.Text = ""
    ckprocmp.Value = 0
    ckprocmod.Value = 0
    ckproccif.Value = 0
    txtmetval.Text = ""
    lblmetval.Caption = ""
    lblidmetval.Caption = 0
    txtfacdist.Text = ""
    txttipdist.Text = ""
    txttipitemdist.Text = ""
    lblfacdist.Caption = ""
    lbltipdist.Caption = ""
    lbltipitemdist.Caption = ""
    lblidfacdist.Caption = 0
    lblidtipdist.Caption = 0
    lblidtipitemdist.Caption = 0
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim xCampos() As String
    Dim nSQLId As String
    Dim xRs As New ADODB.Recordset
    Dim nTitulo As String
    Dim xform As New eps_librerias.FormSeleccion
    Dim A As Integer
    
    Select Case Index
        Case 0 ' Tipo de Distribucion
            ReDim xCampos(2, 4) As String
            
            xCampos(0, 0) = "Codigo":           xCampos(0, 1) = "codigo":           xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Descripción":      xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "3500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
                   
            cSQL = "SELECT mae_metodovaltipdist.id, mae_metodovaltipdist.codigo, mae_metodovaltipdist.descripcion " _
                + vbCr + "FROM mae_metodovaltipdist"
                        
            nTitulo = "Buscando Ítems"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            lblidtipdist.Caption = NulosN(xRs("id"))
            txttipdist.Text = NulosC(xRs("codigo"))
            lbltipdist.Caption = NulosC(xRs("descripcion"))
            txttipitemdist.SetFocus
        
        Case 1 ' Tipo Item Distribucion
            ReDim xCampos(2, 4) As String
            
            xCampos(0, 0) = "Codigo":           xCampos(0, 1) = "codigo":           xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Descripción":      xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "3500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
                   
            cSQL = "SELECT mae_metodovaltipitemdist.id, mae_metodovaltipitemdist.codigo, mae_metodovaltipitemdist.descripcion " _
                + vbCr + "FROM mae_metodovaltipitemdist"
                        
            nTitulo = "Buscando Tipos de Item de Distibucion"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            lblidtipitemdist.Caption = NulosN(xRs("id"))
            txttipitemdist.Text = NulosC(xRs("codigo"))
            lbltipitemdist.Caption = NulosC(xRs("descripcion"))
            txtfacdist.SetFocus
        
        Case 2 ' Factor de Distribucion
            ReDim xCampos(2, 4) As String
            
            xCampos(0, 0) = "Codigo":           xCampos(0, 1) = "codigo":           xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Descripción":      xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "3500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
                   
            cSQL = "SELECT mae_metodovalfacdist.id, mae_metodovalfacdist.codigo, mae_metodovalfacdist.descripcion " _
                + vbCr + "FROM mae_metodovalfacdist"
                        
            nTitulo = "Buscando Factores de Distibucion"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            lblidfacdist.Caption = NulosN(xRs("id"))
            txtfacdist.Text = NulosC(xRs("codigo"))
            lblfacdist.Caption = NulosC(xRs("descripcion"))
        
        Case 3 ' Metodo de Valorizacion
            ReDim xCampos(2, 4) As String
            
            xCampos(0, 0) = "Codigo":           xCampos(0, 1) = "codigo":           xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Descripción":      xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "3500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
                   
            cSQL = "SELECT mae_metodoval.id, mae_metodoval.codigo, mae_metodoval.descripcion " _
                + vbCr + "FROM mae_metodoval"
                        
            nTitulo = "Buscando Metodos de Valorizacion"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            lblidmetval.Caption = NulosN(xRs("id"))
            txtmetval.Text = NulosC(xRs("codigo"))
            lblmetval.Caption = NulosC(xRs("descripcion"))
    End Select
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstDesp
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstDesp.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstDesp("id")), xCon
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        Dim Rpta As Integer
        
        IdMenuActivo = xIdMenu
        
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        SeEjecuto = True
        pCargarGrid
        
    End If
End Sub

Sub Nuevo()
    QueHace = 1
    Blanquea
    Bloquea
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    TituloLabel(1).Caption = "Agregando Metodo de Valorizacion"
    AbrevText.SetFocus
End Sub

Private Sub Form_Load()
    ' CARGAMOS EL FORMULARIO
    QueHace = 3
    SeEjecuto = False
    iniciarCampos
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 3 Then Exit Sub
        MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            RstDesp.Requery
            Dg1.Refresh
            Cancelar
            
            If RstDesp.RecordCount <> 0 Then
                RstDesp.MoveFirst
                RstDesp.Find "id=" & mIdRegistro
                If RstDesp.EOF = True Then RstDesp.MoveFirst
            End If
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstDesp.Filter = ""
    End If
    
    If Button.Index = 15 Then
        Set RstDesp = Nothing
        Unload Me
    End If
End Sub

Function validarDatos() As Boolean
    If NulosC(DescripcionText.Text) = "" Then
        MsgBox "No ha especificado la descripcion del metodo de valorizacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        DescripcionText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosC(AbrevText.Text) = "" Then
        MsgBox "No ha especificado la abreviatura del metodo de valorizacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        AbrevText.SetFocus
        validarDatos = False
        Exit Function
    End If
    
    validarDatos = True
End Function

Function Grabar() As Boolean
    Dim Rpta As Integer
    Dim ConfigVal As New ContabilidadEntidad.EConfigVal
        
On Error GoTo LaCague
    If Not validarDatos Then Grabar = False: Exit Function
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "Grabar", "Modificar") + " el Item", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function
    
    On Error GoTo LaCague
    If QueHace = 1 Then ConfigVal.IdConfiguracionValorizacion = 0 Else ConfigVal.IdConfiguracionValorizacion = NulosN(RstDesp("id"))
        
    ConfigVal.Descripcion = NulosC(DescripcionText.Text)
    ConfigVal.Abreviacion = NulosC(AbrevText.Text)
    ConfigVal.IdMetodoValorizacion = NulosC(lblidmetval.Caption)
    ConfigVal.IdFactorDistribucion = NulosC(lblidfacdist.Caption)
    ConfigVal.IdTipoDistribucion = NulosC(lblidtipdist.Caption)
    ConfigVal.IdTipoItemDistribucion = NulosC(lblidtipitemdist.Caption)
    If ActivoCheck.Value = 0 Then
        ConfigVal.Activo = 0
    Else
        ConfigVal.Activo = -1
    End If
    If ckprocmp.Value = 0 Then
        ConfigVal.ProcesaMP = False
    Else
        ConfigVal.ProcesaMP = True
    End If
    If ckprocmod.Value = 0 Then
        ConfigVal.ProcesaMOD = False
    Else
        ConfigVal.ProcesaMOD = True
    End If
    If ckproccif.Value = 0 Then
        ConfigVal.ProcesaCIF = False
    Else
        ConfigVal.ProcesaCIF = True
    End If
    
    Set ConfigVal.Conexion = xCon
    If Not ConfigVal.Save(0, "") Then Err.Raise &HFFFFFF01, , "Error al intentar grabar el registro"
    
    mIdRegistro = ConfigVal.IdMetodoValorizacion
    MsgBox "El item se grabo con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Set ConfigVal = Nothing
    Exit Function
    
LaCague:
    Grabar = False
    Set ConfigVal = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description), vbCritical, xTitulo
End Function

Sub Cancelar()
    QueHace = 3
    Bloquea
    ActivaTool
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
End Sub

Sub Modificar()
    QueHace = 2
    Bloquea
    ActivaTool
    If TabOne1.CurrTab = 0 Then
        Blanquea
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If

    TabOne1.TabEnabled(0) = False
    TituloLabel(1).Caption = "Modificando Metodo de Valorizacion"
    AbrevText.SetFocus
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    Dim ConfigVal As New ContabilidadEntidad.EConfigVal
    
    ' SI EL ITEM NO TIENE NINGUNA OPERACION SE PROCEDE A ELIMINAR PREVIA AUTORIZACION DEL USUARIO
    Rpta = MsgBox("¿ Esta seguro de eliminar el registro ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        ConfigVal.IdConfiguracionValorizacion = NulosN(RstDesp("id"))
        Set ConfigVal.Conexion = xCon
        ConfigVal.Delete CLng(xIdUsuario), F.MachineName
        
        MsgBox "El registro se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstDesp.Requery
        Dg1.Refresh
        Exit Sub
    End If
End Sub

Sub MuestraSegundoTab()
    Blanquea
    AbrevText.Text = NulosC(RstDesp("abrev"))
    DescripcionText.Text = NulosC(RstDesp("descripcion"))
    txtmetval.Text = NulosC(RstDesp("codmetval"))
    lblmetval.Caption = NulosC(RstDesp("desmetval"))
    lblidmetval.Caption = NulosN(RstDesp("idmetval"))
    
    txtfacdist.Text = NulosC(RstDesp("codfacdist"))
    lblfacdist.Caption = NulosC(RstDesp("desfacdist"))
    lblidfacdist.Caption = NulosN(RstDesp("idfacdist"))
    
    txttipdist.Text = NulosC(RstDesp("codtipdist"))
    lbltipdist.Caption = NulosC(RstDesp("destipdist"))
    lblidtipdist.Caption = NulosN(RstDesp("idtipdist"))
    
    txttipitemdist.Text = NulosC(RstDesp("codtipitemdist"))
    lbltipitemdist.Caption = NulosC(RstDesp("destipitemdist"))
    lblidtipitemdist.Caption = NulosN(RstDesp("idtipitemdist"))
    If RstDesp("activo") Then
        ActivoCheck.Value = 1
    Else
        ActivoCheck.Value = 0
    End If
    If RstDesp("procesamp") Then
        ckprocmp.Value = 1
    Else
        ckprocmp.Value = 0
    End If
    If RstDesp("procesamod") Then
        ckprocmod.Value = 1
    Else
        ckprocmod.Value = 0
    End If
    If RstDesp("procesacif") Then
        ckproccif.Value = 1
    Else
        ckproccif.Value = 0
    End If
End Sub
