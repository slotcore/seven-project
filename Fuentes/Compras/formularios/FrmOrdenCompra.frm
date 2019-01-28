VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOrdenCompra 
   Caption         =   "Compras - Orden de Compra"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   11880
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
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenCompra.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenCompra.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenCompra.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenCompra.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenCompra.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenCompra.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenCompra.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenCompra.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenCompra.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenCompra.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenCompra.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenCompra.frx":2236
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7230
      Left            =   0
      TabIndex        =   13
      Top             =   375
      Width           =   11880
      _cx             =   20955
      _cy             =   12753
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
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
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   45
         TabIndex        =   41
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6300
            Left            =   30
            TabIndex        =   66
            Top             =   480
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11113
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
            Columns(1).Caption=   "Fch.Emi."
            Columns(1).DataField=   "fchemi"
            Columns(1).NumberFormat=   "Short Date"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Tipo Compra"
            Columns(2).DataField=   "desctipcom"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Proveedor"
            Columns(3).DataField=   "nomprov"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Moneda"
            Columns(4).DataField=   "moneda"
            Columns(4).NumberFormat=   "Short Date"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Condición Pago"
            Columns(5).DataField=   "descconpag"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Estado"
            Columns(6).DataField=   "descestado"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1879"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1799"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=131585"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=8043"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=7964"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1826"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1746"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=131588"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=3360"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=3281"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=1746"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1667"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=513"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
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
            _StyleDefs(24)  =   "Splits(0).Style:id=75,.parent=1"
            _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=84,.parent=4"
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=76,.parent=2"
            _StyleDefs(27)  =   "Splits(0).FooterStyle:id=77,.parent=3"
            _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=78,.parent=5"
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=80,.parent=6"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=79,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=81,.parent=8,.bgcolor=&H80&"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=82,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=83,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=85,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=86,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=90,.parent=75"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=76"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=77"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=79"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=98,.parent=75"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=95,.parent=76"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=96,.parent=77"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=97,.parent=79"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=102,.parent=75,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=99,.parent=76"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=100,.parent=77,.alignment=2"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=101,.parent=79"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=110,.parent=75"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=107,.parent=76"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=108,.parent=77"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=109,.parent=79"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=118,.parent=75,.alignment=3"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=115,.parent=76,.alignment=2"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=116,.parent=77,.alignment=2"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=117,.parent=79"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=16,.parent=75"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=13,.parent=76"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=14,.parent=77"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=15,.parent=79"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=75,.alignment=2"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=76"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=77"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=79"
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
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta Orden de Compra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   42
            Top             =   30
            Width           =   11610
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   12525
         TabIndex        =   14
         Top             =   375
         Width           =   11790
         Begin VB.CommandButton cmd 
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   2190
            Picture         =   "FrmOrdenCompra.frx":277E
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   510
            Width           =   240
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   3
            Text            =   "TxtNumSer"
            Top             =   1755
            Width           =   915
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2625
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   4
            Text            =   "TxtNumDoc"
            Top             =   1755
            Width           =   3110
         End
         Begin VB.Frame Frame5 
            Height          =   660
            Left            =   6210
            TabIndex        =   60
            Top             =   450
            Width           =   5310
            Begin VB.Label LblNumero 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "LblNumero"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   2355
               TabIndex        =   62
               Top             =   255
               Width           =   1005
            End
            Begin VB.Label Label2 
               Caption         =   "ORDEN Nº   :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   270
               Left            =   435
               TabIndex        =   61
               Top             =   255
               Width           =   1395
            End
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   2175
            Picture         =   "FrmOrdenCompra.frx":28B0
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   1155
            Width           =   240
         End
         Begin VB.CommandButton CmdBusTipDoc 
            Height          =   240
            Left            =   2175
            Picture         =   "FrmOrdenCompra.frx":29E2
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   1470
            Width           =   240
         End
         Begin VB.CommandButton CmdBusAutoriza 
            Height          =   240
            Left            =   11205
            Picture         =   "FrmOrdenCompra.frx":2B14
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   2805
            Width           =   240
         End
         Begin VB.TextBox TxtNumCot 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   7
            Text            =   "TxtNumCot"
            Top             =   2745
            Width           =   4190
         End
         Begin VB.Frame Frame3 
            Height          =   1080
            Left            =   6210
            TabIndex        =   47
            Top             =   1200
            Width           =   5295
            Begin VB.CommandButton CmdRecha 
               Caption         =   "Rechazar"
               Height          =   315
               Left            =   3285
               TabIndex        =   64
               Top             =   690
               Width           =   1860
            End
            Begin VB.CommandButton CmdAprobada 
               Caption         =   "Aprobar"
               Height          =   315
               Left            =   1395
               TabIndex        =   63
               Top             =   690
               Width           =   1860
            End
            Begin VB.Label LblIdEstado 
               AutoSize        =   -1  'True
               Caption         =   "LblIdEstado"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   255
               TabIndex        =   50
               Top             =   315
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Label LblEstado 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Pendiente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   465
               Left            =   120
               TabIndex        =   48
               Top             =   195
               Width           =   5025
            End
         End
         Begin VB.CommandButton CmdBusContacto 
            Height          =   240
            Left            =   5440
            Picture         =   "FrmOrdenCompra.frx":2C46
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   2460
            Width           =   240
         End
         Begin VB.TextBox TxtContacto 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   6
            Text            =   "TxtContacto"
            Top             =   2430
            Width           =   4190
         End
         Begin VB.CommandButton CmdBusTipoCompra 
            Height          =   240
            Left            =   2175
            Picture         =   "FrmOrdenCompra.frx":2D78
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   840
            Width           =   240
         End
         Begin VB.CommandButton CmdBusCondicion 
            Height          =   240
            Left            =   2175
            Picture         =   "FrmOrdenCompra.frx":2EAA
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   3090
            Width           =   240
         End
         Begin VB.TextBox TxtTipCom 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   0
            Text            =   "TxtTipCom"
            Top             =   810
            Width           =   915
         End
         Begin VB.TextBox TxtConPag 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   8
            Text            =   "TxtConPag"
            Top             =   3060
            Width           =   915
         End
         Begin VB.CommandButton CmdBusProv 
            Height          =   240
            Left            =   2635
            Picture         =   "FrmOrdenCompra.frx":2FDC
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   2145
            Width           =   240
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2460
            Left            =   240
            TabIndex        =   12
            Top             =   3510
            Width           =   11280
            _cx             =   19897
            _cy             =   4339
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
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmOrdenCompra.frx":310E
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmi 
            Height          =   300
            Left            =   7300
            TabIndex        =   10
            Top             =   3090
            Width           =   1260
            _ExtentX        =   2223
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
            Valor           =   "03/02/2007"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchVen 
            Height          =   300
            Left            =   10230
            TabIndex        =   11
            Top             =   3090
            Width           =   1260
            _ExtentX        =   2223
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
            Valor           =   "03/02/2007"
         End
         Begin VB.TextBox TxtNumRuc 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   5
            Text            =   "TxtNumRuc"
            Top             =   2115
            Width           =   1370
         End
         Begin VB.Frame Frame4 
            Height          =   840
            Left            =   240
            TabIndex        =   18
            Top             =   5955
            Width           =   11295
            Begin VB.TextBox TxtTotal 
               Alignment       =   1  'Right Justify
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
               Left            =   10035
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   25
               TabStop         =   0   'False
               Text            =   "TxtTotal"
               Top             =   420
               Width           =   1100
            End
            Begin VB.TextBox TxtIGV 
               Alignment       =   1  'Right Justify
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
               Left            =   7545
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   24
               TabStop         =   0   'False
               Text            =   "TxtIGV"
               Top             =   420
               Width           =   1100
            End
            Begin VB.TextBox TxtBruto 
               Alignment       =   1  'Right Justify
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
               Left            =   5235
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   23
               TabStop         =   0   'False
               Text            =   "TxtBruto"
               Top             =   420
               Width           =   1100
            End
            Begin VB.TextBox TxtInafecto 
               Alignment       =   1  'Right Justify
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
               Left            =   6390
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   22
               TabStop         =   0   'False
               Text            =   "TxtInafect"
               Top             =   420
               Width           =   1100
            End
            Begin VB.CommandButton CmdDelItem 
               Caption         =   "Eliminar Item"
               Enabled         =   0   'False
               Height          =   360
               Left            =   1695
               TabIndex        =   21
               Top             =   285
               Width           =   1335
            End
            Begin VB.CommandButton CmdAddItem 
               Caption         =   "Agregar Item"
               Enabled         =   0   'False
               Height          =   360
               Left            =   270
               TabIndex        =   20
               Top             =   285
               Width           =   1335
            End
            Begin VB.TextBox TxtISC 
               Alignment       =   1  'Right Justify
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
               Left            =   8865
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   19
               TabStop         =   0   'False
               Text            =   "TxtISC"
               Top             =   420
               Width           =   1100
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total"
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
               Index           =   2
               Left            =   10035
               TabIndex        =   31
               Top             =   195
               Width           =   450
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Imp. Afecto"
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
               Index           =   0
               Left            =   5235
               TabIndex        =   29
               Top             =   195
               Width           =   990
            End
            Begin VB.Label LblIgvTasa 
               Alignment       =   2  'Center
               Caption         =   "LblIgvTasa"
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
               Height          =   225
               Left            =   8190
               TabIndex        =   28
               Top             =   195
               Width           =   420
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Inafecto"
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
               Index           =   1
               Left            =   6390
               TabIndex        =   27
               Top             =   195
               Width           =   720
            End
            Begin VB.Line Line4 
               BorderColor     =   &H80000003&
               Index           =   0
               X1              =   3240
               X2              =   3240
               Y1              =   105
               Y2              =   885
            End
            Begin VB.Line Line4 
               BorderColor     =   &H80000005&
               Index           =   1
               X1              =   3255
               X2              =   3255
               Y1              =   90
               Y2              =   870
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "I.S.C."
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
               Index           =   3
               Left            =   8865
               TabIndex        =   26
               Top             =   195
               Width           =   495
            End
            Begin VB.Label LblRotulo 
               AutoSize        =   -1  'True
               Caption         =   "I.G.V. (        ) "
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
               Left            =   7545
               TabIndex        =   30
               Top             =   195
               Width           =   1260
            End
         End
         Begin VB.TextBox TxtTipDoc 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   2
            Text            =   "TxtTipDoc"
            Top             =   1440
            Width           =   915
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   1
            Text            =   "TxtIdMon"
            Top             =   1125
            Width           =   915
         End
         Begin VB.TextBox txtIdAlm 
            Height          =   300
            Left            =   1530
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   68
            Text            =   "txtIdAlm"
            Top             =   480
            Width           =   915
         End
         Begin VB.TextBox TxtAutoriza 
            Height          =   300
            Left            =   7300
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   9
            Text            =   "TxtAutoriza"
            Top             =   2775
            Width           =   4190
         End
         Begin VB.Label LblIdContacto 
            AutoSize        =   -1  'True
            Caption         =   "LblIdContacto"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   5850
            TabIndex        =   71
            Top             =   2490
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label lblAlmacen 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblAlmacen"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   2490
            TabIndex        =   70
            Top             =   480
            Width           =   3230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Almacén"
            Height          =   195
            Index           =   11
            Left            =   255
            TabIndex        =   69
            Top             =   510
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Documento"
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   65
            Top             =   1785
            Width           =   1050
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            Height          =   90
            Left            =   2475
            Top             =   1860
            Width           =   105
         End
         Begin VB.Label LblMoneda 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblMoneda"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   2490
            TabIndex        =   59
            Top             =   1125
            Width           =   3230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   5
            Left            =   255
            TabIndex        =   58
            Top             =   1170
            Width           =   585
         End
         Begin VB.Label LblNomDoc 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblNomDoc"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   2490
            TabIndex        =   56
            Top             =   1440
            Width           =   3230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento"
            Height          =   195
            Index           =   9
            Left            =   255
            TabIndex        =   55
            Top             =   1470
            Width           =   1185
         End
         Begin VB.Label LblIdAutoriza 
            AutoSize        =   -1  'True
            Caption         =   "LblIdAutoriza"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   10530
            TabIndex        =   53
            Top             =   2520
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label LblIdProveedor 
            AutoSize        =   -1  'True
            Caption         =   "LblIdProveedor"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   4350
            TabIndex        =   52
            Top             =   2190
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Cotizacion"
            Height          =   195
            Index           =   8
            Left            =   255
            TabIndex        =   49
            Top             =   2775
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Solicitante"
            Height          =   195
            Index           =   1
            Left            =   6025
            TabIndex        =   45
            Top             =   2805
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Contacto"
            Height          =   195
            Index           =   0
            Left            =   255
            TabIndex        =   44
            Top             =   2460
            Width           =   645
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cond. de Pago"
            Height          =   195
            Index           =   4
            Left            =   255
            TabIndex        =   40
            Top             =   3105
            Width           =   1065
         End
         Begin VB.Label LblNomPro 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblNomPro"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   2925
            TabIndex        =   39
            Top             =   2115
            Width           =   2780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   195
            Index           =   7
            Left            =   255
            TabIndex        =   38
            Top             =   2145
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle Orden de Compra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   345
            Left            =   90
            TabIndex        =   37
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label LblCondPag 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblCondPag"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   2490
            TabIndex        =   36
            Top             =   3060
            Width           =   3230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Emision"
            Height          =   195
            Index           =   2
            Left            =   6030
            TabIndex        =   35
            Top             =   3135
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Vencimiento"
            Height          =   195
            Index           =   3
            Left            =   8895
            TabIndex        =   34
            Top             =   3135
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Item"
            Height          =   195
            Index           =   6
            Left            =   255
            TabIndex        =   33
            Top             =   840
            Width           =   660
         End
         Begin VB.Label LblTipoCompra 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipoCompra"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   2490
            TabIndex        =   32
            Top             =   810
            Width           =   3230
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1058
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
            Object.ToolTipText     =   "Imprimir Guia"
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
   Begin VB.Menu menu1 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Agregar Item              "
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar Item"
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "aprodesa"
      Visible         =   0   'False
      Begin VB.Menu menu2_1 
         Caption         =   "&Aprobar       "
      End
      Begin VB.Menu menu2_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu2_3 
         Caption         =   "&Rechazar"
      End
   End
End
Attribute VB_Name = "FrmOrdenCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMORDENCOMPRA.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO PARA EL REGISTRO DE ORDENES DE COMPRA
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 18/09/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim QueHace  As Integer               ' ESPECIFICA EN QUE MODO SE ENCUENTRA EL FORMULARIO 1 = NUEVO; 2 = MODIFICA; 3 = SOLO LECUTRA
Dim xIdCuenTasa As Integer            ' codigo de la cuenta contable del impuesto
Dim xCuentaDoc As Integer             ' codigo de la cuenta contable del documento
Dim TasaImpuesto As Double            ' ALMACENA LA TASA DEL IMPUESTO
Dim SeEjecuto As Boolean              ' VARIABLE PARA CONTROLAR QUE EL EVENTO CTIVA SE EJECUTE UNA SOLA VEZ
Dim Mostrando As Boolean              ' VARIABLE QUE INFORMA A LOS CONTROLES FlexGrid QUE SE ESTAN AGREGANDO FILAS
Dim CaracteresNumericos As String     ' VARIABLE QUE ALMACENA LOS CARACTERES NUMERICOS QUE SE UTILIZARAN EL LOS CONTROLES TextBox
Dim CaracteresNumericos2 As String    ' VARIABLE QUE ALMACENA LOS CARACTERES NUMERICOS QUE SE UTILIZARAN EL LOS CONTROLES TextBox
Dim xDescImp As String                ' ALMACENA LA DESCRIPCION DEL IMPUESTO
Dim RstTmp As New ADODB.Recordset     ' RECORDSET TEMPORAL
Dim RstOrd As New ADODB.Recordset     ' RECORDSET QUE ALMACENARA Y MOSTRARA LOS REGISTROS DE LA TABLA com_ordencompra
Dim RstTempISC As New ADODB.Recordset ' RECORDSET QUE ALACENARA LOS IMPUESTOS SELECTIVOS

Private Sub CmdAddItem_Click()
    ' AGREGA UNA FILA AL CONTROL FlexGrid Fg1
    If NulosC(Fg1.TextMatrix(Fg1.Rows - 1, 1)) = "" Then Exit Sub
    Fg1.Rows = Fg1.Rows + 1
    
    With Fg1
        .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, 1
    End With
    
    fg1_CellButtonClick Fg1.Rows - 1, 1
    Fg1.SetFocus
End Sub

Private Sub CmdAprobada_Click()
    ' APRUEBA UNA ORDEN DE COMPRA ACTUALIZANDO A 2 EL CAMPO idest DE LA TABLA com_ordencompra
    LblIdEstado.Caption = "2"
    LblEstado.ForeColor = &H8000&
    LblEstado = "Aprobada"
    
    xCon.Execute "UPDATE com_ordencompra SET com_ordencompra.idest = " & Val(LblIdEstado.Caption) & " WHERE (com_ordencompra.id = " & RstOrd("id") & ")"
    RstOrd.Requery
    Dg1.Refresh
End Sub

Private Sub CmdBusAutoriza_Click()
    ' EJECUTA LA BUSQUEDA DE UN PERSONAL
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Apellido Nombre":    xCampos(0, 1) = "apenom":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Cargo":              xCampos(1, 1) = "descar":      xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Codigo":             xCampos(2, 1) = "id":          xCampos(2, 2) = "1000":         xCampos(2, 3) = "N"
    
    xform.SQLCad = "SELECT pla_empleados.id, UCase(pla_empleados.nombre) AS apenom, mae_cargo.descripcion AS descar " _
        & " FROM mae_cargo INNER JOIN pla_empleados ON mae_cargo.id = pla_empleados.idcargo " _
        & " ORDER BY UCase(pla_empleados.nombre)"

    xform.Titulo = "Buscando Personal"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "apenom"
    xform.CampoBusca = "apenom"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtAutoriza.Text = xRs("apenom")
            LblIdAutoriza.Caption = xRs("id")
            TxtFchEmi.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusCondicion_Click()
    ' EJECUTO LA BUSQUEDA DE UNA CONDICION DE PAGO
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_condpago ORDER BY descripcion"
    
    xform.Titulo = "Buscando Condicion de Pago"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtConPag.Text = xRs("id")
            LblCondPag.Caption = xRs("descripcion")
            TxtAutoriza.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusContacto_Click()
    ' EECUTA LA BUSQUEDA DE UN CONTACTO
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(4, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Apellidos y Nombres":   xCampos(0, 1) = "apenom":    xCampos(0, 2) = "3500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Telefono":              xCampos(1, 1) = "numcel":    xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Email":                 xCampos(2, 1) = "email":     xCampos(2, 2) = "2000":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Codigo":                xCampos(3, 1) = "id":        xCampos(3, 2) = "1000":         xCampos(3, 3) = "N"
    
    xform.SQLCad = "SELECT mae_provcontacto.id, UCase([mae_provcontacto]![apecon])+', '+[mae_provcontacto]![nomcon] AS apenom, " _
        & " mae_provcontacto.numcel, mae_provcontacto.email From mae_provcontacto WHERE (((mae_provcontacto.idpro)=" & Val(LblIdProveedor.Caption) & "))"

    xform.Titulo = "Buscando Contactos del Proveedor"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "apenom"
    xform.CampoBusca = "apenom"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtContacto.Text = xRs("apenom")
        LblIdContacto.Caption = xRs("id")
        TxtNumCot.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMon_Click()
    ' EJECUTA A BUSQUEDA DE UNA MONEDA
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_moneda ORDER BY descripcion"
    
    xform.Titulo = "Buscando Moneda"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdMon.Text = xRs("id")
            LblMoneda.Caption = xRs("descripcion")
            TxtTipDoc.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusProv_Click()
    ' EJECUTA LA BUSQUEDA DE UN PROVEEDOR
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Proveedor":    xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":    xCampos(1, 1) = "numruc":      xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_prov.nombre, mae_prov.numruc, mae_prov.id From mae_prov"
    
    xform.Titulo = "Buscando Proveedor"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtNumRuc.Text = xRs("numruc")
            LblNomPro.Caption = xRs("nombre")
            LblIdProveedor.Caption = xRs("id")
            
            Dim xCad As String
            
            xCad = "SELECT mae_provcontacto.id, UCase([mae_provcontacto]![apecon])+', '+[mae_provcontacto]![nomcon] AS apenom, " _
                & " mae_provcontacto.numcel, mae_provcontacto.email From mae_provcontacto " _
                & " WHERE (((mae_provcontacto.idpro)=" & Val(LblIdProveedor.Caption) & ") AND ((mae_provcontacto.defa)=-1))"
    
            Set RstTmp = BuscaConCriterio(xCad, xCon)
    
            If RstTmp.RecordCount <> 0 Then
                RstTmp.MoveFirst
                TxtContacto.Text = RstTmp("apenom")
                LblIdContacto.Caption = RstTmp("id")
            End If
            Set RstTmp = Nothing
            TxtContacto.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipDoc_Click()
    ' EJECUTA LA BUSQUEDA DE UN TIPO DE DOCUMENTO
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, " _
        & " mae_impuestos.Abrev AS abreimp, mae_impuestos.idcuen  as cuentaimp" _
        & " FROM mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id) " _
        & " ON mae_documento.idimp = mae_impuestos.id"
    
    Dim xImpuesto As Double
    
    xform.Titulo = "Buscando Tipo de Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipDoc.Text = xRs("id")
            LblNomDoc.Caption = xRs("descripcion")
            TasaImpuesto = NulosN(xRs("tasa"))
            xDescImp = xRs("descripcion")
            xIdCuenTasa = NulosN(xRs("cuentaimp"))
            LblRotulo = Trim(NulosC(xRs("abreimp"))) + " (       )"
            LblIgvTasa.Caption = Trim(Str(TasaImpuesto)) + "%"
            Frame3.Caption = "( Afecta : " + NulosC(xRs("descimp")) + ")"
            TxtNumSer.SetFocus
            xCuentaDoc = HallaNumCuenta(TxtTipDoc.Text, TxtIdMon.Text)
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipoCompra_Click()
    ' EJECUTA LA BUSQUEDA DE UN TIPO DE PRODUCTO
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_tipoproducto.* FROM mae_tipoproducto"
    
    xform.Titulo = "Buscando Tipo de Producto"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipCom.Text = xRs("id")
            LblTipoCompra = xRs("descripcion")
            TxtIdMon.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdDelItem_Click()
    ' ELIMINA UNA FILA DEL CONTROL Flexgrid Fg1
    If Fg1.Row < 1 Then Exit Sub
    If Fg1.Rows = 1 Then
        MsgBox "No hay items para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Fg1.RemoveItem Fg1.Row
    Fg1.Select 1, 1
    HallarTotal
End Sub

Private Sub CmdRecha_Click()
    ' ACTUALIZA A RECHAZADO UNA ORDEN DE COMPRA, PARA ELLO ACTUALIZA EL CAMPO idest A 4 DE LA TABLA com_ordencompra
    LblIdEstado.Caption = "4"
    LblEstado.ForeColor = &HFF&
    LblEstado = "Rechazada"
    
    xCon.Execute "UPDATE com_ordencompra SET com_ordencompra.idest = " & Val(LblIdEstado.Caption) & " WHERE (com_ordencompra.id = " & RstOrd("id") & ")"
    RstOrd.Requery
    Dg1.Refresh
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
    MuestraSegundoTab
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstOrd
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        Nuevo
    End If
    
    If KeyCode = 46 Then
        Eliminar
    End If
End Sub

Private Sub Dg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu menu2
    End If
End Sub

Private Sub fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    ' EJECUTA LA LA BUSQUEDA DE ITEM EN EL COMBO DE LA COLUMNA 1 DEL CONTROL FlexGrid Fg1
    If Val(TxtTipCom.Text) = 0 Then
        MsgBox "No ha especificado el tipo de compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipCom.SetFocus
        Exit Sub
    End If
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5500":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Unidad":       xCampos(1, 1) = "abrev":          xCampos(1, 2) = "1000":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Codigo":       xCampos(2, 1) = "codpro":         xCampos(2, 2) = "1500":    xCampos(2, 3) = "C"
    
    xform.SQLCad = "SELECT alm_inventario.*, mae_unidades.descripcion AS descuni,  mae_unidades.abrev " _
        & " FROM mae_unidades INNER JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed " _
        & " WHERE tippro = " & Val(TxtTipCom.Text) & " ORDER BY alm_inventario.descripcion"

    xform.Titulo = "Buscando Productos"
    xform.FormaBusca = CualquierParte
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    Mostrando = True
    If xRs.State = 1 Then
        If BuscarItem(xRs("id")) = False Then
            If xRs.RecordCount <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 1) = xRs("descripcion")
                Fg1.TextMatrix(Fg1.Row, 2) = xRs("abrev")
                Fg1.TextMatrix(Fg1.Row, 3) = Format(NulosN(xRs("preuni")), "0.0000")
                Fg1.TextMatrix(Fg1.Row, 6) = xRs("id")
                Fg1.TextMatrix(Fg1.Row, 7) = xRs("idunimed")
                Fg1.TextMatrix(Fg1.Row, 8) = NulosC(xRs("idcuenta"))
                Fg1.TextMatrix(Fg1.Row, 9) = NulosN(xRs("idtipcom"))
            End If
        End If
    End If
    Mostrando = False
    Set xform = Nothing
    Set xRs = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : BuscarItem
'* Tipo             : FUNCION
'* Descripcion      : BUSCA EL ID DE UN ITEM EL CONTROL FlexGrid Fg1
'* Paranetros       : NOMBRE     |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    IdProducto |  INTEGER    |  ESPECIFICA EL ID DEL PRODUCTO
'* Devuelve         :
'*****************************************************************************************************
Function BuscarItem(IdProducto As Integer)
    Dim A As Integer
    
    If Fg1.Rows = 1 Then
        BuscarItem = False
        Exit Function
    End If
    BuscarItem = False
    
    For A = 1 To Fg1.Rows - 1
        If Val(Fg1.TextMatrix(A, 6)) = IdProducto Then
            BuscarItem = True
            MsgBox "El producto seleccionado ya fue agregando a la orden de compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    Next A
End Function

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Mostrando = True Then Exit Sub
    If Col = 4 Or Col = 3 Then
        Fg1.TextMatrix(Fg1.Row, 3) = Format(Fg1.TextMatrix(Fg1.Row, 3), "0.00")
        Fg1.TextMatrix(Fg1.Row, 4) = Format(Fg1.TextMatrix(Fg1.Row, 4), "0.00")
        Fg1.TextMatrix(Fg1.Row, 5) = Val(Fg1.TextMatrix(Fg1.Row, 3)) * Val(Fg1.TextMatrix(Fg1.Row, 4))
        Fg1.TextMatrix(Fg1.Row, 5) = Format(Fg1.TextMatrix(Fg1.Row, 5), "0.0000")
        HallarTotal
        BuscarImpuestos
    End If
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then Exit Sub
    If Fg1.Col = 2 Or Fg1.Col = 5 Then
        Fg1.Editable = flexEDNone
    Else
        Fg1.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = 3 Or Col = 4 Then If InStr(CaracteresNumericos2, Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        CmdAddItem_Click
    End If
    If KeyCode = 46 Then
        CmdDelItem_Click
    End If
    If KeyCode = 93 Then
        PopupMenu menu1
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace = 3 Then Exit Sub
    If Button = 2 Then PopupMenu menu1
End Sub

Private Sub Form_Activate()
    Dim cSQL As String
    ' SEGUNDO EVENTO A EJECUTARSE AL CARGAR EL FORMULARIO
    If SeEjecuto = False Then
        Dim Rpta As Integer
        
        SeEjecuto = True
        
        cSQL = "SELECT mae_prov.nombre AS nomprov, mae_prov.numruc, mae_tipoproducto.descripcion AS desctipcom, mae_estadoordcom.descripcion AS descestado, " _
            & " mae_moneda.simbolo AS moneda, mae_condpago.descripcion AS descconpag, mae_documento.descripcion AS desctipdoc, mae_moneda.descripcion AS descmon, " _
            & " com_ordencompra.*, UCase([pla_empleados]![apepat] & ' ' & [pla_empleados]![apemat]) & ', ' & [pla_empleados]![nom] AS apenomaut " _
            & " FROM mae_moneda RIGHT JOIN (mae_tipoproducto RIGHT JOIN (mae_estadoordcom RIGHT JOIN (mae_prov RIGHT JOIN (((com_ordencompra LEFT JOIN mae_condpago " _
            & " ON com_ordencompra.idconpag = mae_condpago.id) LEFT JOIN mae_documento ON com_ordencompra.idtipdoc = mae_documento.id) LEFT JOIN pla_empleados " _
            & " ON com_ordencompra.idaut = pla_empleados.id) ON mae_prov.id = com_ordencompra.idpro) ON mae_estadoordcom.id = com_ordencompra.idest) " _
            & " ON mae_tipoproducto.id = com_ordencompra.idtippro) ON mae_moneda.id = com_ordencompra.idmon " _
            & vbCr & "ORDER BY com_ordencompra.fchemi DESC"
        RST_Busq RstOrd, cSQL, xCon

        Set Dg1.DataSource = RstOrd
        If RstOrd.RecordCount = 0 Then
            Rpta = MsgBox("No se ha registrado una orden de compra, ¿Desea agregar una ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                Nuevo
            Else
                Set RstOrd = Nothing
                Unload Me
                Exit Sub
            End If
        End If
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
    QueHace = 2
    ActivaTool
    Blanquea
    Bloquea
    Label5.Caption = "Modificando Orden de Compra"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    MuestraSegundoTab
    Frame3.Enabled = False
    Fg1.ColComboList(1) = "|..."
    Fg1.Editable = flexEDKbdMouse
    Fg1.SelectionMode = flexSelectionFree
    TxtTipCom.SetFocus
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
    TxtTipCom.Text = RstOrd("idtippro")
    TxtIdMon.Text = RstOrd("idmon")
    TxtTipDoc.Text = RstOrd("idtipdoc")
    LblIdProveedor.Caption = RstOrd("idpro")
    LblIdContacto.Caption = RstOrd("idcon")
    TxtConPag.Text = RstOrd("idconpag")
    LblIdAutoriza.Caption = RstOrd("idaut")
    TxtFchEmi.Valor = RstOrd("fchemi")
    TxtFchVen.Valor = RstOrd("fchven")
    LblIdEstado.Caption = RstOrd("idest")
    TxtNumCot.Text = RstOrd("numcot")
    TxtNumSer.Text = NulosC(RstOrd("numser"))
    TxtNumDoc.Text = NulosC(RstOrd("numdoc"))
    LblIdAutoriza.Caption = RstOrd("idaut")
    
    LblEstado.Caption = RstOrd("descestado")
    If Val(LblIdEstado.Caption) = 3 Then
        Frame3.Enabled = False
    End If
    
    LblTipoCompra.Caption = RstOrd("desctipcom")
    LblMoneda.Caption = RstOrd("descmon")
    LblNomDoc.Caption = NulosC(RstOrd("desctipdoc"))
    TxtNumRuc.Text = RstOrd("numruc")
    LblNomPro.Caption = RstOrd("nomprov")
    TxtAutoriza.Text = RstOrd("apenomaut")
    LblCondPag.Caption = RstOrd("descconpag")
    
    LblNumero.Caption = NulosC(TxtNumSer.Text) & " - " & NulosC(TxtNumDoc.Text)
    
    If LblIdEstado.Caption = "1" Then LblEstado.Caption = RstOrd("descestado") ': LblEstado.ForeColor = &HC0FFFF
    If LblIdEstado.Caption = "2" Then LblEstado.Caption = RstOrd("descestado") ': LblEstado.ForeColor = &H8000&
    If LblIdEstado.Caption = "3" Then LblEstado.Caption = RstOrd("descestado") ': LblEstado.ForeColor = &HFF0000
    If LblIdEstado.Caption = "4" Then LblEstado.Caption = RstOrd("descestado") ': LblEstado.ForeColor = &HFF&
    
    If NulosN(LblIdContacto.Caption) <> 0 Then
        Set RstTmp = BuscaConCriterio("SELECT * FROM mae_provcontacto WHERE idpro = " & NulosN(LblIdProveedor.Caption) & " And id = " & RstOrd("idcon") & "", xCon)
        If RstTmp.RecordCount <> 0 Then
            TxtContacto.Text = UCase(Trim(RstTmp("nomcon"))) + ", " + Trim(RstTmp("apecon"))
        Else
            TxtContacto.Text = TxtContacto.Text
        End If
        Set RstTmp = Nothing
    End If
    
    Fg1.Rows = 1
    ' Mostramos el detalle de la orden de compra
    Dim xCad As String
    Dim A As Integer
    xCad = "SELECT com_ordencompradet.*, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.idcuenta, " _
        & " alm_inventario.idtipcom FROM mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN com_ordencompradet " _
        & " ON alm_inventario.id = com_ordencompradet.iditem) ON mae_unidades.id = com_ordencompradet.idunimed " _
        & " WHERE (((com_ordencompradet.idcom)=" & RstOrd("id") & "))"

    Set RstTmp = BuscaConCriterio(xCad, xCon)
    Mostrando = True
    
    If RstTmp.RecordCount <> 0 Then
        RstTmp.MoveFirst
        For A = 1 To RstTmp.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = RstTmp("descripcion")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = RstTmp("abrev")
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(RstTmp("preuni"), "0.000000")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(RstTmp("canpro"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(RstTmp("imptot"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = RstTmp("iditem")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = RstTmp("idunimed")
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = RstTmp("idcuenta")
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = RstTmp("idtipcom")
            RstTmp.MoveNext
            If RstTmp.EOF = True Then
                Exit For
            End If
        Next A
    End If
    Mostrando = False
    BuscarImpuestos
    HallarTotal
    If TxtTipDoc.Text <> "" Then LblIgvTasa.Caption = Trim(Str(TasaImpuestoDocumento(Val(TxtTipDoc.Text)))) + "%"
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
    ActivaTool
    Blanquea
    Bloquea
    Label5.Caption = "Agregando Orden de Compra"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    
    Fg1.ColComboList(1) = "|..."
    LblIdEstado.Caption = "1"
    LblEstado.ForeColor = &H8000&
    Frame3.Enabled = False
    LblEstado = "Pendiente"
    Fg1.Rows = 1
    Fg1.Editable = flexEDKbdMouse
    Fg1.SelectionMode = flexSelectionFree
    LblNumero.Caption = Format(HallaCodigoTabla("com_ordencompra", xCon, "id"), "000000")
    TxtTipCom.SetFocus
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    QueHace = 3
    SeEjecuto = False
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    TabOne1.CurrTab = 0
    Fg1.SelectionMode = flexSelectionByRow
    CaracteresNumericos = "0123456789." & Chr(8)
    CaracteresNumericos2 = "0123456789." & Chr(8) & Chr(13)
    LblIgvTasa.Caption = ""
    Fg1.ColWidth(6) = 0
    Fg1.ColWidth(7) = 0
    Fg1.ColWidth(8) = 0
    Fg1.ColWidth(9) = 0
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO, PARA ELLO BLANQUEA
'*                    LOS CONTROLES TextBox DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    TxtTipCom.Text = ""
    TxtNumRuc.Text = ""
    TxtContacto.Text = ""
    TxtNumCot.Text = ""
    TxtIdMon.Text = ""
    TxtConPag.Text = ""
    TxtAutoriza.Text = ""
    TxtFchEmi.Valor = ""
    TxtFchVen.Valor = ""
    TxtTipDoc.Text = ""
    
    LblTipoCompra.Caption = ""
    LblIdProveedor.Caption = ""
    LblNomPro.Caption = ""
    LblMoneda.Caption = ""
    LblCondPag.Caption = ""
    LblIdContacto.Caption = ""
    LblIdAutoriza.Caption = ""
    LblNomDoc.Caption = ""
    TxtBruto.Text = ""
    TxtInafecto.Text = ""
    TxtIGV.Text = ""
    TxtISC.Text = ""
    TxtTotal.Text = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
End Sub

'*****************************************************************************************************
'* Nombre           : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES TextBox DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Bloquea()
    TxtTipCom.Locked = Not TxtTipCom.Locked
    TxtNumRuc.Locked = Not TxtNumRuc.Locked
    TxtTipDoc.Locked = Not TxtTipDoc.Locked
    TxtNumCot.Locked = Not TxtNumCot.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    TxtConPag.Locked = Not TxtConPag.Locked
    TxtFchEmi.Locked = Not TxtFchEmi.Locked
    TxtFchVen.Locked = Not TxtFchVen.Locked
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    
    CmdAddItem.Enabled = Not CmdAddItem.Enabled
    CmdDelItem.Enabled = Not CmdDelItem.Enabled
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LA BARRA DE HERRAMIENTAS DEL FORMULARIO
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

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario mientras este ingresando o modificando una orden de compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 2
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub Menu1_1_Click()
    CmdAddItem_Click
End Sub

Private Sub Menu1_3_Click()
    CmdDelItem_Click
End Sub

Private Sub menu2_1_Click()
    If RstOrd("idest") = 3 Then
        MsgBox "Esta orden ya ha sido procesada, no se puede modificar su estado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    xCon.Execute "UPDATE com_ordencompra SET com_ordencompra.idest = 2 WHERE (com_ordencompra.id = " & RstOrd("id") & ")"
    RstOrd.Requery
    Dg1.Refresh
End Sub

Private Sub menu2_3_Click()
    If RstOrd("idest") = 3 Then
        MsgBox "Esta orden ya ha sido procesada, no se puede modificar su estado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    xCon.Execute "UPDATE com_ordencompra SET com_ordencompra.idest = 4 WHERE (com_ordencompra.id = " & RstOrd("id") & ")"
    RstOrd.Requery
    Dg1.Refresh
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then
        If RstOrd("idest") = 3 Then
            MsgBox "No puede modificar una orden de compra procesada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Modificar
    End If
    If Button.Index = 3 Then
        If RstOrd("idest") = 3 Then
            MsgBox "No puede eliminar una orden de compra procesada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Eliminar
    End If
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstOrd.Requery
            Dg1.Refresh
            Dg1.SetFocus
        End If
    End If
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 10 Then Buscar
        
    If Button.Index = 12 Then ImprimirOrden
    
    If Button.Index = 14 Then
        Set RstOrd = Nothing
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : ImprimirOrden
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPRIME LA ORDEN DE COMPRA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ImprimirOrden()
    TabOne1.CurrTab = 0
    
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT com_ordencompradet.*, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.idcuenta, " _
        & " alm_inventario.idtipcom FROM mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN com_ordencompradet " _
        & " ON alm_inventario.id = com_ordencompradet.iditem) ON mae_unidades.id = com_ordencompradet.idunimed  " _
        & " WHERE (((com_ordencompradet.idcom)=" & RstOrd("id") & "))", xCon
    
    Set RptOrdenCompra.DataSource = Rst
    
    With RptOrdenCompra.Sections("Sección2")
        .Controls("TxtNomEmp").Caption = Trim(NomEmp)
        .Controls("TxtNumRuc").Caption = Trim(NumRUC)
        .Controls("TxtFecha").Caption = Format(Date, "dd/mm/yyyy")
        .Controls("TxtNumOrd").Caption = Format(RstOrd("id"), "000000")
        .Controls("TxtTipoCom").Caption = RstOrd("desctipcom")
        .Controls("TxtMoneda").Caption = RstOrd("moneda")
        .Controls("TxtProv").Caption = RstOrd("nomprov")
        .Controls("TxtNumCot").Caption = RstOrd("numcot")
        .Controls("TxtCondPag").Caption = RstOrd("descconpag")
        .Controls("TxtSolicitante").Caption = ""
        .Controls("TxtAutoriza").Caption = ""
        .Controls("TxtFchEmi").Caption = RstOrd("fchemi")
        .Controls("TxtFchFin").Caption = RstOrd("fchven")
    End With
    RptOrdenCompra.Width = 12000
    RptOrdenCompra.Height = 8010
    RptOrdenCompra.Show
    
'    ' Exportar Excel Reporte
'    RptOrdenCompra.ExportReport rptKeyHTML, "c:\report.html"
'    Dim Abrir_Excel As Object
'    Set Abrir_Excel = CreateObject("Excel.Application")
'    Abrir_Excel.Visible = True
'    Abrir_Excel.Workbooks.Open ("c:\report.html")
'    Abrir_Excel.Windows("report.html").Activate
'    Abrir_Excel.Sheets("report").Select
End Sub

'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA LA BUSQUEDA DE UN REGISTRO EN EL RECORDSET RstOrd
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Buscar()
    TabOne1.CurrTab = 0
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Proveedor":         xCampos(0, 1) = "nomprov":          xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº Ruc":            xCampos(1, 1) = "numruc":           xCampos(1, 2) = "1500":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Nº Orden Compra":   xCampos(2, 1) = "id":               xCampos(2, 2) = "1680":         xCampos(2, 3) = "N"
    
    xform.SQLCad = "SELECT mae_prov.nombre AS nomprov, mae_prov.numruc, mae_trabajadores.apellnom, com_ordencompra.id " _
        & " FROM mae_trabajadores RIGHT JOIN (mae_prov RIGHT JOIN com_ordencompra ON mae_prov.id = com_ordencompra.idpro) " _
        & " ON mae_trabajadores.id = com_ordencompra.idaut ORDER BY mae_trabajadores.apellnom, com_ordencompra.id"

    xform.Titulo = "Buscando Orden de Compra"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nomprov"
    xform.CampoBusca = "nomprov"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        RstOrd.MoveFirst
        RstOrd.Find "id = " & xRs("id") & ""
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL PROCESO DE INGRESAR O MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    ActivaTool
    Bloquea
    Label5.Caption = "Detalle Orden de Compra"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Fg1.Editable = flexEDNone
    Frame3.Enabled = True
    Fg1.SelectionMode = flexSelectionByRow
    Dg1.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA com_ordencompra
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    
    Rpta = MsgBox("¿Esta seguro de eliminar la orden de compra seleccionada?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        TabOne1.CurrTab = 0
        xCon.Execute "DELETE * FROM com_ordencompra WHERE id = " & RstOrd("id") & ""
        MsgBox "La orden de compra se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstOrd.Requery
        Dg1.Refresh
        Dg1.SetFocus
    End If
End Sub

Private Sub TxtAutoriza_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtAutoriza_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusAutoriza_Click
    End If
End Sub

Private Sub TxtConPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If NulosC(TxtConPag.Text) <> "" Then
            Set RstTmp = BuscaConCriterio("SELECT * FROM mae_condpago WHERE id = " & Val(TxtConPag.Text) & "", xCon)
            
            If RstTmp.RecordCount <> 0 Then
                LblCondPag.Caption = RstTmp("descripcion")
            Else
                TxtConPag.Text = ""
                LblCondPag.Caption = ""
            End If
        End If
        Set RstTmp = Nothing
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtConPag_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCondicion_Click
    End If
End Sub

Private Sub TxtContacto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtContacto_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusContacto_Click
    End If
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If NulosC(TxtIdMon.Text) <> "" Then
            Set RstTmp = BuscaConCriterio("SELECT * FROM mae_moneda WHERE id = " & Val(TxtIdMon.Text) & "", xCon)
            
            If RstTmp.RecordCount <> 0 Then
                LblMoneda.Caption = RstTmp("descripcion")
            Else
                TxtIdMon.Text = ""
                LblMoneda.Caption = ""
            End If
        End If
        Set RstTmp = Nothing
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

Private Sub TxtNumCot_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtNumDoc.Text) <> "" Then
        If IsNumeric(TxtNumDoc.Text) = True Then
            TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
        End If
        If NulosC(TxtNumDoc.Text) <> "" And NulosC(TxtNumSer.Text) <> "" Then
            If ExisteNumDocOrdenCompra = True Then
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub TxtNumRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtNumRuc.Text <> "" Then
            Dim xRs1 As New ADODB.Recordset
            RST_Busq xRs1, "SELECT * FROM mae_prov WHERE numruc like '" & TxtNumRuc.Text & "%' ORDER BY numruc", xCon
            If xRs1.RecordCount <> 0 Then
                TxtNumRuc.Text = xRs1("numruc")
                LblNomPro.Caption = xRs1("nombre")
                LblIdProveedor.Caption = xRs1("id")
                Set xRs1 = Nothing
                
                Dim xCad As String
            
                xCad = "SELECT mae_provcontacto.id, UCase([mae_provcontacto]![apecon])+', '+[mae_provcontacto]![nomcon] AS apenom, " _
                    & " mae_provcontacto.numcel, mae_provcontacto.email From mae_provcontacto " _
                    & " WHERE (((mae_provcontacto.idpro)=" & Val(LblIdProveedor.Caption) & ") AND ((mae_provcontacto.defa)=-1))"
        
                Set RstTmp = BuscaConCriterio(xCad, xCon)
        
                If RstTmp.RecordCount <> 0 Then
                    RstTmp.MoveFirst
                    TxtContacto.Text = RstTmp("apenom")
                    LblIdContacto.Caption = RstTmp("id")
                End If
                Set RstTmp = Nothing
            Else
                LblIdProveedor.Caption = ""
                LblNomPro.Caption = ""
                TxtNumRuc.Text = ""
            End If
        Else
            TxtNumRuc.Text = ""
            LblNomPro.Caption = ""
            LblIdProveedor.Caption = ""
        End If
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumRuc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusProv_Click
    End If
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumSer_Validate(Cancel As Boolean)
    Dim RstNumOrden As New ADODB.Recordset
    
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtNumSer.Text) <> "" Then
        TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
        RST_Busq RstNumOrden, "SELECT com_ordencompra.numser, com_ordencompra.numdoc From com_ordencompra Where (((com_ordencompra.numser) = '" & Format(TxtNumSer.Text, "0000") & "'))" _
            & " ORDER BY com_ordencompra.numdoc", xCon

        If RstNumOrden.RecordCount = 0 Then
            TxtNumDoc.Text = "0000000001"
        Else
            RstNumOrden.MoveLast
            TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
            TxtNumDoc.Text = Format((Val(RstNumOrden("numdoc")) + 1), "0000000000")
        End If
    End If
    Set RstNumOrden = Nothing
End Sub

Private Sub TxtTipCom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If NulosC(TxtTipCom.Text) <> "" Then
            Set RstTmp = BuscaConCriterio("SELECT * FROM mae_tipoproducto WHERE id = " & Val(TxtTipCom.Text) & "", xCon)
            
            If RstTmp.RecordCount <> 0 Then
                LblTipoCompra.Caption = RstTmp("descripcion")
            Else
                TxtTipCom.Text = ""
                LblTipoCompra.Caption = ""
            End If
        End If
        Set RstTmp = Nothing
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipCom_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipoCompra_Click
    End If
End Sub

Function ExisteNumDocOrdenCompra() As Boolean
    Dim Rst As New ADODB.Recordset
    Dim nSQL As String
    If QueHace <> 1 Then nSQL = " and com_ordencompra.id <> " & NulosN(RstOrd("id"))
    
    ' BUSCAMOS EN NUMERO DE DOCUMENTO
    RST_Busq Rst, "SELECT com_ordencompra.fchemi, com_ordencompra.numcot, com_ordencompra.numser & '-' & com_ordencompra.numdoc As numdoc FROM com_ordencompra WHERE numser = '" & NulosC(TxtNumSer.Text) & "' and numdoc = '" & NulosC(TxtNumDoc.Text) & "'" & nSQL, xCon
    If Rst.RecordCount = 0 Then
        ' SI NO EXISTE ESTA BIEN
        ExisteNumDocOrdenCompra = False
    Else
        ' SI EXISTE ESTA MAL
        MsgBox "El número de documento ingresado ya fue registrado" & vbCr & "Nº Documento: " & NulosC(Rst("numdoc")) & vbCr & "Fecha Doc.   " & NulosC(Rst("fchemi")) & vbCr & "Ingrese Otro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDoc.Text = ""

        ExisteNumDocOrdenCompra = True
    End If
    Set Rst = Nothing
End Function

'*****************************************************************************************************
'* Nombre           : HallarTotal
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : HALLA EL TOTAL DE LOS ITEMS DEL CONTROL FlexGrid Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub HallarTotal()
    Dim A As Integer
    Dim Total, TotalIna As Double
    
    For A = 1 To Fg1.Rows - 1
        If Val(Fg1.TextMatrix(A, 9)) = "4" Then
            TotalIna = TotalIna + Val(Fg1.TextMatrix(A, 5))
        Else
            Total = Total + Val(Fg1.TextMatrix(A, 5))
        End If
    Next A
    
    TxtBruto.Text = Format(NulosN(Total), "0.00")
    TxtInafecto.Text = Format(TotalIna, "0.00")
    TxtIGV.Text = NulosN(Total) * ((TasaImpuesto / 100) + 1)
    TxtIGV.Text = Val(TxtIGV.Text) - Val(TxtBruto.Text)
    TxtTotal.Text = ((Val(TxtBruto.Text) + (Val(TxtInafecto.Text) + Val(TxtIGV.Text))) + Val(TxtISC.Text))
    TxtTotal.Text = Format(TxtTotal.Text, "0.00")
    TxtIGV.Text = Format(TxtIGV.Text, "0.00")
End Sub

'*****************************************************************************************************
'* Nombre           : BuscarImpuestos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : BUSCA LOS IMPUESTOS AMARRADOS AL ITEM SELECCIONADO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub BuscarImpuestos()
    Dim A As Integer
    Dim xImpSEL, xImpIGV As Double
    
    Dim Rst As New ADODB.Recordset
    
    Set RstTempISC = Nothing
    PreparaRST_ISC
    xImpSEL = 0
    ' buscando selectivo
    For A = 1 To Fg1.Rows - 1
        RST_Busq Rst, "SELECT mae_impuestos.tasa, mae_impuestos.idcuen, con_planctas.cuenta " _
            & " FROM (alm_inventario LEFT JOIN mae_impuestos ON alm_inventario.idimpsel = mae_impuestos.id) " _
            & " LEFT JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id WHERE " _
            & " (((alm_inventario.id) = val(" & Fg1.TextMatrix(A, 6) & " )))", xCon
        
        If NulosN(Rst("idcuen")) <> 0 Then
            xImpSEL = xImpSEL + Val(Fg1.TextMatrix(A, 5)) * (Rst("tasa") / 100)
            
            If RstTempISC.RecordCount = 0 Then
                RstTempISC.AddNew
                RstTempISC("idcuen") = Rst("idcuen")
                RstTempISC("total") = RstTempISC("total") + Val(Fg1.TextMatrix(A, 5)) * (Rst("tasa") / 100)
            Else
                RstTempISC.MoveFirst
                RstTempISC.Find "idcuen = " & Rst("idcuen") & ""
                
                If RstTempISC.EOF = False Then
                    RstTempISC("idcuen") = Rst("idcuen")
                    RstTempISC("total") = RstTempISC("total") + Val(Fg1.TextMatrix(A, 5)) * (Rst("tasa") / 100)
                End If
            End If
        End If
    Next A
    
    TxtISC.Text = Format(NulosN(xImpSEL), "0.00")
    
    ' buscando el impuesto a las ventas
    xImpIGV = 0
    For A = 1 To Fg1.Rows - 1
        If Fg1.TextMatrix(A, 9) <> 4 Then
            xImpIGV = xImpIGV + Fg1.TextMatrix(A, 5) * (Val(LblIgvTasa.Caption) / 100)
        End If
    Next A
    
    TxtIGV.Text = Format(xImpIGV, "0.00")
End Sub

'*****************************************************************************************************
'* Nombre           : PreparaRST_ISC
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL RECORDSET TEMPORAL PARA ALAMCENAR LOS DATOS DEL IMPUESTO SELECTIVO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub PreparaRST_ISC()
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(9, 3) As String

    xCampos(0, 0) = "idcuen":        xCampos(0, 1) = "N":      xCampos(0, 2) = "2"
    xCampos(1, 0) = "Total":         xCampos(1, 1) = "D":      xCampos(1, 2) = "2"
    Set RstTempISC = xFun.CrearRstTMP(xCampos)

    RstTempISC.Open
End Sub

Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        
        If NulosC(TxtTipDoc.Text) = "" Then Exit Sub
        Dim xRs As New ADODB.Recordset
        
        RST_Busq xRs, "SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, " _
        & " mae_impuestos.Abrev AS abreimp, mae_impuestos.idcuen as cuentaimp " _
        & " FROM mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuen = con_planctas.id) " _
        & " ON mae_documento.idimp = mae_impuestos.id WHERE mae_documento.id  = " & Val(TxtTipDoc.Text) & "", xCon
        
        If xRs.RecordCount = 0 Then
            TxtTipDoc.Text = ""
            LblNomDoc.Caption = ""
        Else
            TxtTipDoc.Text = xRs("id")
            LblNomDoc.Caption = xRs("descripcion")
            TasaImpuesto = NulosN(xRs("tasa"))
            xDescImp = xRs("descripcion")
            xIdCuenTasa = NulosN(xRs("cuentaimp"))
            LblRotulo = Trim(NulosC(xRs("abreimp"))) + " (       )"
            LblIgvTasa.Caption = Trim(Str(TasaImpuesto)) + "%"
        End If
        Set xRs = Nothing
        xCuentaDoc = HallaNumCuenta(TxtTipDoc.Text, TxtIdMon.Text)
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipDoc_Click
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : FUNCION
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA com_ordencompra, ESTA FUNCION DEVUELVE VERDADERO
'*                    CUANDO TIENE EXITO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function Grabar() As Boolean
    ' VERIFICAMOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
    If NulosC(TxtTipCom.Text) = "" Then
        MsgBox "No ha especificado el tipo de compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipCom.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtIdMon.Text) = "" Then
        MsgBox "No ha especificado la moneda", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMon.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtTipDoc.Text) = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtNumRuc.Text) = "" Then
        MsgBox "No ha especificado proveedor para la orden de compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumRuc.SetFocus
        Exit Function
    End If
    
'    If NulosC(TxtNumCot.Text) = "" Then
'        MsgBox "No ha especificado el numero de la cotizacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtNumCot.SetFocus
'        Exit Function
'    End If
    
    If NulosC(TxtConPag.Text) = "" Then
        MsgBox "No ha especificado la condicion de pago", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtConPag.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtAutoriza.Text) = "" Then
        MsgBox "No ha especificado el usuario que autoriza la Orden de Compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtAutoriza.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtFchEmi.Valor) = "" Then
        MsgBox "No ha especificado la fecha de emision de la Orden de Compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchEmi.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtFchVen.Valor) = "" Then
        MsgBox "No ha especificado la fecha de vencimiento de la Orden de Compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchVen.SetFocus
        Exit Function
    End If
    
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado items para la orden de compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        CmdAddItem.SetFocus
        Exit Function
    End If
    
    If CDate(TxtFchEmi.Valor) > CDate(TxtFchVen.Valor) Then
        MsgBox "La fecha de emision no puede ser mayor a la fecha de vencimiento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchEmi.SetFocus
        Exit Function
    End If
    
    Dim A As Integer
    
    ' VERIFICAMOS QUE LOS DATOS DEL DETALLE SEAN LOS CORRECTOS
    For A = 1 To Fg1.Rows - 1
        If NulosC(Fg1.TextMatrix(A, 1)) = "" Then
            MsgBox "No ha especificado el producto en el item Nº " + Trim(Str(A)) + " ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    
        If NulosN(Fg1.TextMatrix(A, 3)) = 0 Then
            MsgBox "No ha especificado el precio unitario del item Nº " + Trim(Str(A)) + " ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    
        If NulosN(Fg1.TextMatrix(A, 4)) = 0 Then
            MsgBox "No ha especificado la cantidad del producto en el item Nº " + Trim(Str(A)) + " ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    Next A
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xId As Integer
    
    On Error GoTo LaCague
    
    xCon.BeginTrans
    
    
    If QueHace = 1 Then
        ' SI ES UNA ORDEN DE COMPRA NUEVO
        ' OBTENEMOS EL ULTIMO ID DE LA TABLA com_ordencompra
        xId = HallaCodigoTabla("com_ordencompra", xCon, "id")
        LblNumero.Caption = Format(xId, "000000")
        
        RST_Busq RstCab, "SELECT * FROM com_ordencompra", xCon
        RST_Busq RstDet, "SELECT * FROM com_ordencompradet", xCon
        
        RstCab.AddNew
        RstCab("id") = xId
    Else
        ' SI SE ESTA MODIFICANDO UN REGISTRO
        xId = RstOrd("id")
        RST_Busq RstCab, "SELECT * FROM com_ordencompra WHERE id = " & xId & "", xCon
        xCon.Execute "DELETE * FROM com_ordencompradet WHERE idcom = " & xId & ""
        
        RST_Busq RstDet, "SELECT * FROM com_ordencompradet", xCon
    End If
    
    RstCab("idtippro") = Val(TxtTipCom.Text)
    RstCab("idpro") = Val(LblIdProveedor.Caption)
    If Val(LblIdContacto.Caption) <> 0 Then RstCab("idcon") = Val(LblIdContacto.Caption)
    RstCab("idmon") = Val(TxtIdMon.Text)
    RstCab("idconpag") = Val(TxtConPag.Text)
    RstCab("numcot") = NulosC(TxtNumCot.Text)
    RstCab("numser") = NulosC(TxtNumSer.Text)
    RstCab("numdoc") = NulosC(TxtNumDoc.Text)
    RstCab("idaut") = NulosN(LblIdAutoriza.Caption)
    RstCab("fchemi") = TxtFchEmi.Valor
    RstCab("fchven") = TxtFchVen.Valor
    RstCab("idtipdoc") = Val(TxtTipDoc.Text)
    RstCab("idest") = 1
    RstCab.Update
    
    ' GRABAMOS EL DETALLE DEL REGISTRO
    For A = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("idcom") = xId
        RstDet("iditem") = Val(Fg1.TextMatrix(A, 6))
        RstDet("idunimed") = Val(Fg1.TextMatrix(A, 7))
        RstDet("preuni") = Val(Fg1.TextMatrix(A, 3))
        RstDet("canpro") = Val(Fg1.TextMatrix(A, 4))
        RstDet("imptot") = Val(Fg1.TextMatrix(A, 5))
        RstDet.Update
    Next A
    
    xCon.CommitTrans
    
    MsgBox "La orden de compra se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function
