VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManTareas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Producción - Tareas"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   8820
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
            Picture         =   "FrmManTareas.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTareas.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTareas.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTareas.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTareas.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTareas.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTareas.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTareas.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTareas.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTareas.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTareas.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTareas.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTareas.frx":277E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   0
      TabIndex        =   6
      Top             =   375
      Width           =   8805
      _cx             =   15531
      _cy             =   12726
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
         Height          =   6795
         Left            =   9450
         TabIndex        =   10
         Top             =   375
         Width           =   8715
         Begin VB.Frame Frame3 
            Height          =   5175
            Left            =   600
            TabIndex        =   11
            Top             =   480
            Width           =   7530
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   3
               Left            =   1755
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   2
               Text            =   "txt(3)"
               Top             =   2025
               Width           =   2115
            End
            Begin VB.TextBox txt 
               BackColor       =   &H0080FF80&
               Height          =   315
               Index           =   0
               Left            =   6120
               TabIndex        =   21
               Tag             =   "null"
               Text            =   "txt(0)"
               Top             =   255
               Visible         =   0   'False
               Width           =   1170
            End
            Begin VB.CommandButton cb 
               Height          =   240
               Index           =   0
               Left            =   2265
               Picture         =   "FrmManTareas.frx":2B10
               Style           =   1  'Graphical
               TabIndex        =   17
               ToolTipText     =   "Seleccione la Unidad de Medida"
               Top             =   2385
               Width           =   240
            End
            Begin VB.Frame fra 
               Caption         =   "¿Es Tarea Diversa?"
               Enabled         =   0   'False
               Height          =   885
               Index           =   0
               Left            =   735
               TabIndex        =   16
               Top             =   3150
               Width           =   6000
               Begin VB.OptionButton opt_diverso 
                  Caption         =   "Si"
                  Height          =   195
                  Index           =   1
                  Left            =   3420
                  TabIndex        =   5
                  Top             =   360
                  Width           =   480
               End
               Begin VB.OptionButton opt_diverso 
                  Caption         =   "No"
                  Height          =   195
                  Index           =   0
                  Left            =   1200
                  TabIndex        =   4
                  Top             =   375
                  Value           =   -1  'True
                  Width           =   555
               End
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   1
               Left            =   1755
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   0
               Text            =   "txt(1)"
               Top             =   1380
               Width           =   1455
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   2
               Left            =   1755
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   1
               Text            =   "txt(2)"
               Top             =   1695
               Width           =   4875
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   0
               Left            =   1755
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   3
               Text            =   "txt_cb(0)"
               ToolTipText     =   "Ingrese la Unidad de Medida"
               Top             =   2355
               Width           =   780
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "Abrev"
               Height          =   195
               Index           =   3
               Left            =   735
               TabIndex        =   23
               Top             =   2130
               Width           =   420
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "codigo"
               Height          =   195
               Index           =   0
               Left            =   5565
               TabIndex        =   22
               Top             =   375
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Unidad"
               Height          =   195
               Index           =   0
               Left            =   735
               TabIndex        =   20
               Top             =   2415
               Width           =   510
            End
            Begin VB.Label lbl_cb_cod 
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "lbl_cb_cod(0)"
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
               Left            =   5355
               TabIndex        =   18
               Top             =   2355
               Visible         =   0   'False
               Width           =   1290
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "Descripción"
               Height          =   195
               Index           =   2
               Left            =   735
               TabIndex        =   13
               Top             =   1800
               Width           =   840
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "Código"
               Height          =   195
               Index           =   1
               Left            =   735
               TabIndex        =   12
               Top             =   1485
               Width           =   495
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
               Left            =   2535
               TabIndex        =   19
               Top             =   2355
               Width           =   2730
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Tarea"
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
            TabIndex        =   14
            Top             =   30
            Width           =   8610
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6795
         Left            =   45
         TabIndex        =   7
         Top             =   375
         Width           =   8715
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   30
            TabIndex        =   8
            Top             =   345
            Width           =   8640
            _ExtentX        =   15240
            _ExtentY        =   11324
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "ID"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Código"
            Columns(1).DataField=   "codigo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descripción"
            Columns(2).DataField=   "descripcion"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "U.M."
            Columns(3).DataField=   "abrev"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   16
            Columns(4)._MaxComboItems=   5
            Columns(4).ValueItems(0)._DefaultItem=   0
            Columns(4).ValueItems(0).Value=   "0"
            Columns(4).ValueItems(0).Value.vt=   8
            Columns(4).ValueItems(0).DisplayValue=   "No"
            Columns(4).ValueItems(0).DisplayValue.vt=   8
            Columns(4).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
            Columns(4).ValueItems(1)._DefaultItem=   0
            Columns(4).ValueItems(1).Value=   "-1"
            Columns(4).ValueItems(1).Value.vt=   8
            Columns(4).ValueItems(1).DisplayValue=   "Si"
            Columns(4).ValueItems(1).DisplayValue.vt=   8
            Columns(4).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
            Columns(4).ValueItems.Count=   2
            Columns(4).Caption=   "Diverso"
            Columns(4).DataField=   "diverso"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2143"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2064"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=7594"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=7514"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1138"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1058"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1482"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1402"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
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
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Tareas"
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
            TabIndex        =   9
            Top             =   30
            Width           =   8580
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8820
      _ExtentX        =   15558
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
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
Attribute VB_Name = "FrmManTareas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMMANTAREAS.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : PERMITE EL INGRESO DE ALTA Y BAJAS DE LAS TAREAS QUE UTILIZARA EL SISTEMA
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 06/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstFrm As New ADODB.Recordset   ' RECORDSET QUE ALAMCENARA LOS DATOS DE LA TABLA pro_tareas
Dim QueHace As Integer              ' INDICA EN QUE MODO SE ENCUENTRA EL FORMULARIO
Private SeEjecuto As Boolean        ' VERIFICA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim mIdRegistro&                    ' identificador del registro
Dim fOrdenLista As Boolean          ' especfica el orden de la lista de la consulta
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO

'*****************************************************************************************************
'* Nombre           : Filtrar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA UN FILTRO SOBRE EL RECORDSET RstFrm
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Filtrar()
    Dim xCampos(3, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Código":           xCampos(0, 1) = "codigo":        xCampos(0, 2) = "C":         xCampos(0, 3) = "4200"
    xCampos(1, 0) = "Descripcion":      xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "C":         xCampos(1, 3) = "1500"
    xCampos(2, 0) = "Abrev":            xCampos(2, 1) = "abrev":         xCampos(2, 2) = "C":         xCampos(2, 3) = "1500"
    
    TabOne1.CurrTab = 0
    
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstFrm
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    'ORDENA EN FORMA ASCENDENTE O DESENDENTE LAS COLUMNAS DEL CONTROL Dg1
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

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDOS E CARGUE EL FORMULARIO
    On Error GoTo error
    If SeEjecuto = False Then

        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        SeEjecuto = True
        
        RST_Busq RstFrm, "SELECT pro_tareas.id, pro_tareas.codigo,pro_tareas.abrev as nomcorto, pro_tareas.descripcion, pro_tareas.diverso, pro_tareas.idunimed, mae_unidades.abrev " _
            + vbCr + " FROM mae_unidades RIGHT JOIN pro_tareas ON mae_unidades.id = pro_tareas.idunimed " _
            + vbCr + " ORDER BY pro_tareas.descripcion; ", xCon

        Set Dg1.DataSource = RstFrm

    End If
    Exit Sub

error:
    SHOW_ERROR Me.Name, "Form_Activate"
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
    Label5.Caption = "Agregando Tarea"
    
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    txt(1).Text = pCargarCodigo()
    txt(2).SetFocus
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
'* Descripcion      : PREPRA LOS CONTROLES TEXTBOX PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    LimpiaText txt
    LimpiaText txt_cb
    LimpiaText lbl_cb_cod
    LimpiaText txt_cb
End Sub

'*****************************************************************************************************
'* Nombre           : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES TEXTBOX
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Bloquea(band As Boolean)
    habilitar_Locked txt, band
    
    habilitar_Locked txt_cb, band
    
    fra(0).Enabled = Not band
End Sub

Private Sub Form_Load()
    'PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    CentrarFrm Me
    QueHace = 3
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO ACTUAL
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    Blanquea
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then Exit Sub
    txt(1).Text = NulosC(RstFrm("codigo"))
    txt(2).Text = NulosC(RstFrm("descripcion"))
    txt(3).Text = NulosC(RstFrm("nomcorto"))
    
    If NulosN(RstFrm("diverso")) = -1 Then
        opt_diverso(1).Value = True
    Else
        opt_diverso(0).Value = True
    End If
    
    ' de la unidad
    If NulosN(RstFrm("idunimed")) <> 0 Then
        txt_cb(0).Text = RstFrm("idunimed")
        txt_cb_Validate 0, False
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL INGRESO O EDICION DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    Bloquea True
    ActivaTool
    Label5.Caption = "Detalle de la Tarea"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Dg1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario mientras este ingresando o modificando una tarea", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    Else
        Set RstFrm = Nothing
        SeEjecuto = False
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
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
        TDB_FiltroLimpiar Dg1
        RstFrm.Filter = adFilterNone
        RstFrm.Requery
        Dg1.Refresh
    End If
    
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 12 Then pExportar
    
    If Button.Index = 15 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA pro_tareas, ESTA FUNCION DEVUELVE VERDADERO CUANDO
'*                    TIENE EXITO
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Function Grabar() As Boolean
    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " la Tarea", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    Dim RstCab As New ADODB.Recordset
    Dim xId As Double

On Error GoTo LaCague
    xCon.BeginTrans
    If QueHace = 1 Then
        xId = HallaCodigoTabla("pro_tareas", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM pro_tareas", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstFrm("id")
        RST_Busq RstCab, "SELECT * FROM pro_tareas WHERE id = " & RstFrm("id") & "", xCon
    End If
    
    mIdRegistro = xId
    
    RstCab("codigo") = pCargarCodigo()
    RstCab("descripcion") = Trim(txt(2).Text)
    RstCab("abrev") = Trim(txt(3).Text)
    RstCab("idunimed") = NulosN(lbl_cb_cod(0).Caption)
    
    If opt_diverso(0).Value = True Then
        RstCab("diverso") = 0
    Else
        RstCab("diverso") = -1
    End If
    
    RstCab.Update
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    
    Set RstCab = Nothing:
    xCon.CommitTrans
    
    MsgBox "La Tarea se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    
    Grabar = True
    Exit Function
        
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing:
    MsgBox "No se pudo guardar la Tarea por el siguiente motivo: " + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
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
    Label5.Caption = "Modificando Tarea"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    MuestraSegundoTab
    txt(2).SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA pro_receta
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    Dim RstBusca  As New ADODB.Recordset
    Dim nSQL As String
    Dim xId&
    
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        If RstFrm.RecordCount = 0 Then
            MsgBox "No hay registros", vbExclamation, xTitulo
        Else
            MsgBox "Seleccione un Registro para Eliminar", vbExclamation, xTitulo
        End If
        Exit Sub
    End If
    
    TabOne1.CurrTab = 0
    
    xId = NulosN(RstFrm.Fields("id"))
    
    nSQL = "SELECT TOP 1 'Receta' AS Origen, alm_inventario.descripcion as nombre " _
    + vbCr + " FROM (alm_inventario INNER JOIN pro_receta ON alm_inventario.id = pro_receta.iditem) INNER JOIN (pro_tareas INNER JOIN pro_recetatar ON pro_tareas.id = pro_recetatar.idtar) ON pro_receta.id = pro_recetatar.idrec " _
    + vbCr + " WHERE pro_recetatar.idtar = " & xId & " ; " _
    + vbCr + " UNION " _
    + vbCr + " SELECT TOP 1 'Programación Diaria' AS origen, 'Fecha Producción: ' & Format([pro_progdia].[fchprod],'dd/mm/yy') AS nombre " _
    + vbCr + " FROM (pro_progdia INNER JOIN pro_progdiadet ON pro_progdia.id = pro_progdiadet.idprogra) INNER JOIN pro_progdiadettar ON pro_progdiadet.idprogra = pro_progdiadettar.idprogra " _
    + vbCr + " WHERE pro_progdiadettar.idtar= " & xId & " ; " _
    + vbCr + " UNION " _
    + vbCr + " SELECT TOP 1 'Control de Tareas ' AS origen, 'Fecha Trabajo : ' & Format([pro_controltar].[fchtra],'dd/mm/yy') & '    Area:  ' & [mae_area].[descripcion] AS nombre " _
    + vbCr + " FROM (mae_area INNER JOIN pro_controltar ON mae_area.id = pro_controltar.idarea) INNER JOIN pro_controltardet ON pro_controltar.id = pro_controltardet.idctr " _
    + vbCr + " WHERE pro_controltardet.idtar = " & xId & " ; "

    ' si el registro tiene relaciones mostrara un menaje
    RST_Busq RstBusca, nSQL, xCon
    If RstBusca.EOF = False Or RstBusca.BOF = False Or RstBusca.RecordCount <> 0 Then
        MsgBox "El registro no se puede eliminar" + vbCr + "Esta asociado a " & RstBusca("origen") & vbCr & RstBusca("nombre"), vbExclamation, xTitulo
        Set RstBusca = Nothing
        Exit Sub
    End If
    Set RstBusca = Nothing
    
    Rpta = MsgBox("Esta seguro de eliminar la Tarea seleccionada", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM pro_tareas WHERE id =" & RstFrm("id") & " "
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstFrm("id") & " AND idform = " & IdMenuActivo

        MsgBox "La Tarea se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
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
    Dim xRs As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Código":           xCampos(0, 1) = "codigo":        xCampos(0, 2) = "1200":         xCampos(0, 3) = "c"
    xCampos(1, 0) = "Descripcion":      xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "4500":        xCampos(1, 3) = "c"
    xCampos(2, 0) = "Abrev":            xCampos(2, 1) = "abrev":         xCampos(2, 2) = "800":         xCampos(2, 3) = "c"
    
    TabOne1.CurrTab = 0
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, RstFrm.Source, xCampos(), "Buscando Tareas", "descripcion", "descripcion", Principio
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    
    RstFrm.MoveFirst
    RstFrm.Find "id = " & xRs("id") & ""

SALIR:
    Set xRs = Nothing

error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub


'*****************************************************************************************************
'* Nombre           : fValidarDatos
'* Tipo             : FUNCION
'* Descripcion      : validar que no haya duplicidad de datos, DEVUELVE VERDADERO SI TODO ESTA CORRECTO
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Private Function fValidarDatos() As Boolean
    If NulosC(txt(2).Text) = "" Then
        MsgBox "No ha especificado la descripción de la tarea", vbInformation, xTitulo
        txt(2).SetFocus
        Exit Function
    End If
    
    Dim RstTmp As New ADODB.Recordset
    If QueHace = 1 Then
        RST_Busq RstTmp, "SELECT descripcion FROM pro_tareas WHERE ucase(descripcion)='" + UCase(Trim(txt(2).Text)) + "';", xCon
    Else
        RST_Busq RstTmp, "SELECT descripcion FROM pro_tareas WHERE ucase(descripcion)='" + UCase(Trim(txt(2).Text)) + "' AND id <> " + CStr(RstFrm.Fields("id")) + ";", xCon
    End If
    
    If RstTmp.BOF = False Or RstTmp.EOF = False Or RstTmp.RecordCount <> 0 Then
        MsgBox "La Tarea " + IIf(QueHace = 1, " ya fue ingresado", "ya existe"), vbExclamation, xTitulo
        Set RstTmp = Nothing
        Exit Function
    End If
    
    Set RstTmp = Nothing
    fValidarDatos = True
End Function
 
Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cb_cod(Index).Caption = ""
    End If
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    
    If txt_cb(Index).Locked = True Then Exit Sub
    
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)
    If txt_cb(Index).Text = "" Then Exit Sub
    Dim RST_TMP As New ADODB.Recordset
    Dim nSQL As String

    nSQL = "SELECT mae_unidades.id, mae_unidades.descripcion AS nombre, mae_unidades.id AS cod, mae_unidades.abrev " _
        + vbCr + " FROM mae_unidades " _
        + vbCr + " WHERE mae_unidades.id=" & NulosN(txt_cb(Index).Text) & "; "
            
    If xCon.State = 0 Then Exit Sub
    RST_Busq RST_TMP, nSQL, xCon
    
    If RST_TMP.State = 0 Then Exit Sub
    
    If RST_TMP.RecordCount > 0 Then
        txt_cb(Index).Text = RST_TMP.Fields(0) & ""         ' TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RST_TMP.Fields(1) & ""      ' NOMBRE
        lbl_cb_cod(Index).Caption = RST_TMP.Fields(2) & ""  ' CODIGO
    Else
        txt_cb(Index).Text = ""
    End If
    
    Set RST_TMP = Nothing
    Exit Sub

error:
    Set RST_TMP = Nothing
    SHOW_ERROR Me.Name, "txt_cb_KeyDown(" + CStr(Index) + ")"
End Sub

Private Sub cb_Click(Index As Integer)
    If QueHace = 3 Then Exit Sub
    
    Dim xRs As New ADODB.Recordset
    Dim xCampos() As String
    Dim nSQL As String
    
    On Error GoTo error
    
    nSQL = "SELECT mae_unidades.id, mae_unidades.descripcion AS nombre, mae_unidades.id AS cod, mae_unidades.abrev " _
        + vbCr + " FROM mae_unidades;"
   
    ReDim xCampos(2, 3) As String
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "nombre":   xCampos(0, 2) = "3500":  xCampos(0, 3) = "C"
    xCampos(1, 0) = "Abrev":        xCampos(1, 1) = "abrev":    xCampos(1, 2) = "600":   xCampos(1, 3) = "C"
    xCampos(2, 0) = "Id":           xCampos(2, 1) = "id":       xCampos(2, 2) = "500":   xCampos(1, 3) = "N"
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Unidades", "nombre", "nombre", Principio
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    txt_cb(Index) = xRs.Fields(0) & ""             ' TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & ""     ' NOMBRE
    lbl_cb_cod(Index).Caption = xRs.Fields(2) & "" ' CODIGO

SALIR:
    Set xRs = Nothing
    Exit Sub

error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Function pCargarCodigo() As String
    Dim Rst As New ADODB.Recordset
    Dim mCodigo As Long
    RST_Busq Rst, "select top 1 pro_tareas.codigo from pro_tareas order by codigo desc", xCon
    
    If Rst.RecordCount <> 0 Then
        mCodigo = NulosN(Mid(Rst("codigo"), 6)) + 1
    Else
        mCodigo = 1
    End If
    
    Set Rst = Nothing
    pCargarCodigo = "TAREA" & Format(mCodigo, "0000")
End Function

'*****************************************************************************************************
'* Nombre           : pExportar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA A MS EXCEL LOS DATOS DEL RECORDSET RstTmp
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pExportar()
    TabOne1.CurrTab = 0
        
    Dim xCampos(4, 3) As String
    Dim oExport As New SGI2_funciones.formularios
    Dim RstTmp As New ADODB.Recordset
    Set RstTmp = RstFrm.Clone
    '0 = Nombre a Mostrar;
    '1 = nombre de Campo del Rst;
    '2 = alineacion(0::derecha, 1::centro, 2::izquierda);
    '3 = ancho de columna
    ' obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Código":       xCampos(0, 1) = "codigo":       xCampos(0, 2) = 0:  xCampos(0, 3) = "1200"
    xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = 0:  xCampos(1, 3) = "3900"
    xCampos(2, 0) = "Abrev":        xCampos(2, 1) = "nomcorto":     xCampos(2, 2) = 0:  xCampos(2, 3) = "2400"
    xCampos(3, 0) = "Unidad":       xCampos(3, 1) = "abrev":        xCampos(3, 2) = 0:  xCampos(3, 3) = "750"
    xCampos(4, 0) = "Es Diverso":   xCampos(4, 1) = "diverso":      xCampos(4, 2) = 0:  xCampos(4, 3) = "800"
    
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "Lista de Tareas", "", "", "Listado de Tareas", RstTmp, xCampos()
    
    Set oExport = Nothing
    Set RstTmp = Nothing
End Sub
