VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManPuntosVenta 
   Caption         =   "Ventas - Mantenimiento de Puntos de Venta"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11730
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   11730
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6810
      Left            =   15
      TabIndex        =   4
      Top             =   375
      Width           =   11715
      _cx             =   20664
      _cy             =   12012
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
         Height          =   6390
         Left            =   12360
         TabIndex        =   8
         Top             =   375
         Width           =   11625
         Begin VB.Frame Frame3 
            Height          =   3045
            Left            =   1110
            TabIndex        =   10
            Top             =   1710
            Width           =   9465
            Begin VB.CommandButton CmdBusDep 
               Enabled         =   0   'False
               Height          =   240
               Left            =   7380
               Picture         =   "FrmManPuntosVenta.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   2220
               Width           =   240
            End
            Begin VB.CommandButton CmdBusDis 
               Enabled         =   0   'False
               Height          =   240
               Left            =   7380
               Picture         =   "FrmManPuntosVenta.frx":0132
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   1905
               Width           =   240
            End
            Begin VB.TextBox txtDepartamento 
               Height          =   300
               Left            =   2730
               TabIndex        =   20
               Text            =   "TxtDepartamento"
               Top             =   2190
               Width           =   4920
            End
            Begin VB.TextBox txtDistrito 
               Height          =   300
               Left            =   2730
               TabIndex        =   18
               Text            =   "TxtDistrito"
               Top             =   1875
               Width           =   4920
            End
            Begin VB.TextBox TxtCodPunVen 
               Height          =   300
               Left            =   2730
               Locked          =   -1  'True
               MaxLength       =   13
               TabIndex        =   1
               Text            =   "TxtCodPunVen"
               Top             =   930
               Width           =   1980
            End
            Begin VB.TextBox TxtDirPunVen 
               Height          =   300
               Left            =   2730
               Locked          =   -1  'True
               MaxLength       =   60
               TabIndex        =   3
               Text            =   "TxtDirPunVen"
               Top             =   1560
               Width           =   4920
            End
            Begin VB.CommandButton CmdBusCli 
               Enabled         =   0   'False
               Height          =   240
               Left            =   7380
               Picture         =   "FrmManPuntosVenta.frx":0264
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   645
               Width           =   240
            End
            Begin VB.TextBox TxtNomPunVen 
               Height          =   300
               Left            =   2730
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   2
               Text            =   "TxtNomPunVen"
               Top             =   1245
               Width           =   4920
            End
            Begin VB.TextBox TxtCliente 
               Height          =   300
               Left            =   2730
               Locked          =   -1  'True
               TabIndex        =   0
               Text            =   "TxtCliente"
               Top             =   615
               Width           =   4920
            End
            Begin VB.Label LblIdDepartamento 
               AutoSize        =   -1  'True
               Caption         =   "LblIdDepartamento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   7725
               TabIndex        =   24
               Top             =   2220
               Visible         =   0   'False
               Width           =   1620
            End
            Begin VB.Label LblIdDistrito 
               AutoSize        =   -1  'True
               Caption         =   "LblIdDistrito"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   7755
               TabIndex        =   23
               Top             =   1950
               Visible         =   0   'False
               Width           =   1035
            End
            Begin VB.Label lbldepartamento 
               AutoSize        =   -1  'True
               Caption         =   "Departamento"
               Height          =   195
               Left            =   1080
               TabIndex        =   19
               Top             =   2205
               Width           =   1005
            End
            Begin VB.Label lbldistrito 
               AutoSize        =   -1  'True
               Caption         =   "Distrito"
               Height          =   195
               Left            =   1080
               TabIndex        =   17
               Top             =   1890
               Width           =   480
            End
            Begin VB.Label LblIdCliente 
               AutoSize        =   -1  'True
               Caption         =   "LblIdCliente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   7755
               TabIndex        =   16
               Top             =   660
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Codigo Punto Venta"
               Height          =   195
               Left            =   1080
               TabIndex        =   15
               Top             =   975
               Width           =   1425
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Direccion"
               Height          =   195
               Left            =   1080
               TabIndex        =   14
               Top             =   1575
               Width           =   675
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Nombre Punto Venta"
               Height          =   195
               Left            =   1080
               TabIndex        =   13
               Top             =   1290
               Width           =   1485
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Cliente"
               Height          =   195
               Left            =   1080
               TabIndex        =   12
               Top             =   660
               Width           =   480
            End
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle Punto de Venta"
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
            Left            =   15
            TabIndex        =   9
            Top             =   30
            Width           =   11520
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6390
         Left            =   45
         TabIndex        =   5
         Top             =   375
         Width           =   11625
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   5970
            Left            =   60
            TabIndex        =   6
            Top             =   375
            Width           =   11535
            _ExtentX        =   20346
            _ExtentY        =   10530
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
            Columns(1).Caption=   "Cliente"
            Columns(1).DataField=   "nombre"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Cod. Punto Venta"
            Columns(2).DataField=   "codcen"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Nombre "
            Columns(3).DataField=   "descripcion"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Dirección"
            Columns(4).DataField=   "dir"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=5027"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=4948"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=2884"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2805"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=5821"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=5741"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=5609"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=5530"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=512"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000018&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
            _StyleDefs(11)  =   ":id=2,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H80&"
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
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=0"
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
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Consulta Puntos de Venta"
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
            TabIndex        =   7
            Top             =   30
            Width           =   11520
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7755
      Top             =   45
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
            Picture         =   "FrmManPuntosVenta.frx":0396
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPuntosVenta.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPuntosVenta.frx":0C6C
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPuntosVenta.frx":0DF0
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPuntosVenta.frx":1244
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPuntosVenta.frx":135C
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPuntosVenta.frx":18A0
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPuntosVenta.frx":1DE4
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPuntosVenta.frx":1EF8
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPuntosVenta.frx":200C
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPuntosVenta.frx":2460
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPuntosVenta.frx":25CC
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPuntosVenta.frx":2B14
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManPuntosVenta.frx":2E2E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
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
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir Documento"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Documento"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Exportar a Excel"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmManPuntosVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre            : FrmManPuntosVenta
'* Tipo              : FORMULARIO
'* Descripcion       : FORMULARIO QUE PERMITE EL REGISTRO DE LOS PUNTOS DE VENTA DE CADA CLIENTE
'* DISEÑADO POR      : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION   : 24/09/09
'* VERSION           : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstPunto As New ADODB.Recordset    ' RECORDSET QUE ALMACENARA LOS REGISTROS DE LA TABLA VTA_PuntoVenta, LOS DATOS SE MOSTRARAN EN LA PESTAÑA Consulta
Dim QueHace As Integer                 ' VARIABLE QUE INDICA EN QUE MODO SE ENCUENTRA EL FORMULARIO 1 = ADICIONA, 2 = MODIFICA, 3 = SOLOLECTURA
Dim SeEjecuto As Boolean               ' VARIABLE QUE VERIFICA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim mIdRegistro& '--identificador del registro
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


Private Sub CmdBusCli_Click()
    ' EJECUTA LA BUSQUEDA DE UN CLIENTE
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre":     xCampos(0, 1) = "nombre":      xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":     xCampos(1, 1) = "id":          xCampos(1, 2) = "1200":    xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT MAE_Cliente.id, MAE_Cliente.Nombre From MAE_Cliente ORDER BY MAE_Cliente.Nombre"
    
    xform.Titulo = "Buscando Clientes"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtCliente.Text = xRs("nombre")
        LblIdCliente.Caption = xRs("id")
        TxtCodPunVen.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusDep_Click()
    ' EJECUTA LA BUSQUEDA DE UN DEPARTAMENTO
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Departamento":     xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":           xCampos(1, 1) = "id":           xCampos(1, 2) = "1200":    xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_departamentos.id, mae_departamentos.descripcion From mae_departamentos ORDER BY mae_departamentos.descripcion"
    
    xform.Titulo = "Buscando Distritos"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        txtDepartamento.Text = xRs("descripcion")
        LblIdDepartamento.Caption = xRs("id")
        TxtCodPunVen.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusDis_Click()
    ' EJECUTA LA BUSQUEDA DE UN DISTRITO
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Distrito":     xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":       xCampos(1, 1) = "id":           xCampos(1, 2) = "1200":    xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_distritos.id, mae_distritos.descripcion From mae_distritos ORDER BY mae_distritos.descripcion"
    
    xform.Titulo = "Buscando Distritos"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        txtDistrito.Text = xRs("descripcion")
        LblIdDistrito.Caption = xRs("id")
        TxtCodPunVen.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Dg1_Click()
    'Dg1.ExportRows
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstPunto("id")), xCon
    End If
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARUIO
    If SeEjecuto = False Then
'        Dim Rpta As Integer
        SeEjecuto = True
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon

        ' CARGAMOS LOS DATOS DE LA TABLA VTA_PuntoVenta AL RECORSET RstPunto
        RST_Busq RstPunto, " SELECT VTA_PuntoVenta.*, MAE_Cliente.nombre, [mae_distritos.descripcion] AS Distrito, [mae_departamentos.descripcion] AS Departamento " & _
            " FROM ((VTA_PuntoVenta LEFT JOIN MAE_Cliente ON VTA_PuntoVenta.idcli = MAE_Cliente.id) LEFT JOIN mae_distritos ON VTA_PuntoVenta.iddis = mae_distritos.id) LEFT JOIN mae_departamentos ON VTA_PuntoVenta.iddep = mae_departamentos.id " & _
            " ORDER BY VTA_PuntoVenta.idcli, VTA_PuntoVenta.descripcion ", xCon

        Set Dg1.DataSource = RstPunto
'        If RstPunto.RecordCount = 0 Then
'            Rpta = MsgBox("No se ha registrado ningun punto de ventas, ¿desea agregar uno ahora?", vbYesNo + vbQuestion + vbDefaultButton1, xTitulo)
'            If Rpta = vbYes Then
'                Nuevo
'            Else
'                Set RstPunto = Nothing
'                Unload Me
'            End If
'        End If

    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO, BLANQUEA LOS CONTROLES
'*                    TextBox
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    TxtCliente.Text = ""
    TxtCodPunVen.Text = ""
    TxtNomPunVen.Text = ""
    TxtDirPunVen.Text = ""
    txtDepartamento = ""
    txtDistrito = ""
End Sub

'*****************************************************************************************************
'* Nombre           : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESCATIVA LOS CONTROLES TextBox
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Bloquea()
    TxtCliente.Locked = Not TxtCliente.Locked
    TxtCodPunVen.Locked = Not TxtCodPunVen.Locked
    TxtNomPunVen.Locked = Not TxtNomPunVen.Locked
    TxtDirPunVen.Locked = Not TxtDirPunVen.Locked
    CmdBusCli.Enabled = Not CmdBusCli.Enabled
    CmdBusDis.Enabled = Not CmdBusDis.Enabled
    CmdBusDep.Enabled = Not CmdBusDep.Enabled
    txtDistrito.Locked = Not txtDistrito.Locked
    txtDepartamento.Locked = Not txtDepartamento.Locked
    
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
    xHorIni = Time
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label1.Caption = "Agregando Punto de Venta"
    Bloquea
    Blanquea
    TxtCliente.SetFocus
End Sub

Private Sub Form_Load()
    ' SEGUNDO EVENTO A EJECUTARSE AL CARGAR EL FORMULARIO
    SeEjecuto = False
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    TabOne1.CurrTab = 0
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO, ESTA INFORMACION SE VISUALIZA EN LA
'*                    PESTAÑA DETALLE
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    TxtCliente.Text = NulosC(RstPunto("nombre"))
    LblIdCliente.Caption = RstPunto("idcli")
    TxtCodPunVen.Text = RstPunto("codcen")
    TxtNomPunVen.Text = RstPunto("descripcion")
    TxtDirPunVen.Text = RstPunto("dir")
    LblIdDepartamento.Caption = RstPunto("iddep")
    LblIdDistrito.Caption = RstPunto("iddis")
    txtDistrito.Text = NulosC(RstPunto("distrito"))
    txtDepartamento.Text = NulosC(RstPunto("departamento"))
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstPunto.Requery
            Dg1.Refresh
            '--------------------------------------------------------------------------
            If RstPunto.RecordCount <> 0 Then
                RstPunto.MoveFirst
                RstPunto.Find "id=" & mIdRegistro
                If RstPunto.EOF = True Then RstPunto.MoveFirst
            End If
            '--------------------------------------------------------------------------
            
            
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 13 Then
        Dim xFun As New eps_librerias.FuncionesDGrid
        xFun.xNomEmp = NomEmp
        xFun.xNumRuc = NumRUC
        xFun.ExportarDGExcel RstPunto, Dg1, "PUNTOS DE VENTA POR CLIENTE"
        Set xFun = Nothing
    End If
    
    If Button.Index = 16 Then
        Set RstPunto = Nothing
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA UNA BUSQUEDA DE CLIENTE
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Buscar()
'    Dim xform As New eps_librerias.FormBuscar
'    Dim xRs As New ADODB.Recordset
'    Dim xCampos(2, 4) As String
'
'    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
'    xCampos(0, 0) = "Nombre":     xCampos(0, 1) = "nombre":      xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
'    xCampos(1, 0) = "Codigo":     xCampos(1, 1) = "id":          xCampos(1, 2) = "1200":    xCampos(1, 3) = "N"
'
'    xform.SQLCad = "SELECT MAE_Cliente.id, MAE_Cliente.Nombre From MAE_Cliente ORDER BY MAE_Cliente.Nombre"
'
'    xform.Titulo = "Buscando Clientes"
'    xform.FormaBusca = Principio
'    xform.Criterio = ""
'    xform.Ordenado = "nombre"
'    xform.CampoBusca = "nombre"
'    Set xform.Coneccion = xCon
'    Set xRs = xform.BuscarReg(xCampos)
'    If xRs.State = 1 Then
'        TxtCliente.Text = xRs("nombre")
'        LblIdCliente.Caption = xRs("id")
'        TxtCodPunVen.SetFocus
'    End If
'    Set xform = Nothing
'    Set xRs = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA Y DESACTIVA LA BARRA DE HERRAMIENTAS
'* Paranetros       :
'* Devuelve         :
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
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL INGRESO O MODIFICACION DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    Label1.Caption = "Detalle Punto de Venta"
    Bloquea
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    QueHace = 3
    ActivaTool
    Dg1.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    QueHace = 2
    xHorIni = Time
    ActivaTool
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label1.Caption = "Modificando Punto de Venta"
    TxtCliente.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PERMITE ELIMINAR UN REGISTRO DE LA TABLA VTA_PuntoVenta
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    
    Rpta = MsgBox("Esta seguro de eliminar el punto de venta " + RstPunto("descripcion"), vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM VTA_PuntoVenta WHERE id = " & RstPunto("id") & ""
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstPunto("id") & " AND idform = " & IdMenuActivo
        
        RstPunto.Requery
        Dg1.Refresh
        MsgBox "El punto de venta se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : FUNCION
'* Descripcion      : GRABAR UN REGISTRO EN LA TABLA VTA_PuntoVenta, ESTA FUNCION DEVUELVE VERDADERO
'*                    SI TIENE EXITO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function Grabar() As Boolean
    ' VERIFICAMOS QUE LOS DATOS NECESARIOS ESTEN CORRECTAMENTE INGRESADOS
    If TxtCliente.Text = "" Then
        MsgBox "No ha especificado el nombre del cliente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtCliente.SetFocus
        Exit Function
    End If
    
    If TxtCodPunVen.Text = "" Then
        MsgBox "No ha especificado el codigo del punto de venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtCodPunVen.SetFocus
        Exit Function
    End If
    
    If TxtNomPunVen.Text = "" Then
        MsgBox "No ha especificado el nombre del punto de venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNomPunVen.SetFocus
        Exit Function
    End If
    
    If TxtDirPunVen.Text = "" Then
        MsgBox "No ha especificado la direccion del punto de venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtDirPunVen.SetFocus
        Exit Function
    End If
    
    Dim RstGra As New ADODB.Recordset
    Dim xId As Double
    
On Error GoTo LaCague
    xCon.BeginTrans
    
    If QueHace = 1 Then
        ' SI ES UN NUEVO REGISTRO
        xId = HallaCodigoTabla("VTA_PuntoVenta", xCon, "id")     ' OBTENEMOS EL NUEVO ID DEL REGISTRO
        RST_Busq RstGra, "SELECT * FROM VTA_PuntoVenta", xCon
        RstGra.AddNew
        RstGra("id") = xId
    Else
        xId = RstPunto("id")
        RST_Busq RstGra, "SELECT * FROM VTA_PuntoVenta WHERE id = " & xId & "", xCon
    End If
    
    mIdRegistro = xId
    
    ' GRABAMOS LOS DATOS DEL REGISTRO
    RstGra("idcli") = NulosN(LblIdCliente.Caption)
    RstGra("codcen") = TxtCodPunVen.Text
    RstGra("descripcion") = TxtNomPunVen.Text
    RstGra("dir") = TxtDirPunVen.Text
    RstGra("iddis") = NulosN(LblIdDistrito)
    RstGra("iddep") = NulosN(LblIdDepartamento)
    RstGra.Update
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    xCon.CommitTrans
    MsgBox "El punto de venta se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    Set RstGra = Nothing
    Grabar = True
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    Set RstGra = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

Private Sub TxtCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtCliente_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCli_Click
    End If
End Sub

Private Sub TxtCodPunVen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtDirPunVen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNomPunVen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub
