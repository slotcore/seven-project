VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmManUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SEVEN - Mantenimiento de Usuarios"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   Icon            =   "FrmManUsuarios.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9330
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
            Picture         =   "FrmManUsuarios.frx":030A
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarios.frx":084E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarios.frx":0BE0
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarios.frx":0D64
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarios.frx":11B8
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarios.frx":12D0
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarios.frx":1814
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarios.frx":1D58
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarios.frx":1E6C
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarios.frx":1F80
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarios.frx":23D4
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManUsuarios.frx":2540
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   4980
      Left            =   0
      TabIndex        =   7
      Top             =   375
      Width           =   9330
      _cx             =   16457
      _cy             =   8784
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
         Height          =   4560
         Left            =   9975
         TabIndex        =   12
         Top             =   375
         Width           =   9240
         Begin VB.Frame Frame3 
            Height          =   3240
            Left            =   435
            TabIndex        =   13
            Top             =   840
            Width           =   8460
            Begin VB.CommandButton CmdBusTipuUser 
               Height          =   240
               Left            =   2700
               Picture         =   "FrmManUsuarios.frx":2A88
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   1185
               Width           =   240
            End
            Begin VB.TextBox TxtCorreo 
               Height          =   300
               Left            =   2130
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   3
               Text            =   "TxtCorreo"
               Top             =   1485
               Width           =   4000
            End
            Begin VB.TextBox TxtTipUser 
               Height          =   300
               Left            =   2130
               Locked          =   -1  'True
               MaxLength       =   1
               TabIndex        =   2
               Text            =   "TxtTipUser"
               Top             =   1155
               Width           =   840
            End
            Begin VB.TextBox TxtApe 
               Height          =   300
               Left            =   2130
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   0
               Text            =   "TxtApe"
               Top             =   525
               Width           =   4000
            End
            Begin VB.TextBox TxtLogin 
               Height          =   300
               Left            =   2130
               Locked          =   -1  'True
               MaxLength       =   15
               TabIndex        =   4
               Text            =   "TxtLogin"
               Top             =   1860
               Width           =   1500
            End
            Begin VB.TextBox TxtContra 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   2130
               Locked          =   -1  'True
               MaxLength       =   15
               PasswordChar    =   "*"
               TabIndex        =   5
               Text            =   "TxtContra"
               Top             =   2175
               Width           =   1500
            End
            Begin VB.TextBox TxtContra2 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   2130
               Locked          =   -1  'True
               MaxLength       =   15
               PasswordChar    =   "*"
               TabIndex        =   6
               Text            =   "TxtContra2"
               Top             =   2490
               Width           =   1500
            End
            Begin VB.TextBox TxtNom 
               Height          =   300
               Left            =   2130
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   1
               Text            =   "TxtNom"
               Top             =   840
               Width           =   4000
            End
            Begin VB.Label LblTipoUser 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblTipoUser"
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
               Left            =   3030
               TabIndex        =   24
               Top             =   1155
               Width           =   3105
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "E-mail"
               Height          =   195
               Index           =   3
               Left            =   375
               TabIndex        =   22
               Top             =   1515
               Width           =   420
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Usuario"
               Height          =   195
               Index           =   1
               Left            =   375
               TabIndex        =   21
               Top             =   1170
               Width           =   1125
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Apellidos"
               Height          =   195
               Index           =   10
               Left            =   375
               TabIndex        =   18
               Top             =   555
               Width           =   630
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Confirmar Contraseña"
               Height          =   195
               Index           =   8
               Left            =   375
               TabIndex        =   17
               Top             =   2520
               Width           =   1515
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Usuario"
               Height          =   195
               Index           =   6
               Left            =   375
               TabIndex        =   16
               Top             =   1875
               Width           =   540
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Contraseña"
               Height          =   195
               Index           =   2
               Left            =   375
               TabIndex        =   15
               Top             =   2205
               Width           =   810
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Nombre"
               Height          =   195
               Index           =   0
               Left            =   375
               TabIndex        =   14
               Top             =   870
               Width           =   555
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Usuario"
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
            TabIndex        =   19
            Top             =   30
            Width           =   9105
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4560
         Left            =   45
         TabIndex        =   8
         Top             =   375
         Width           =   9240
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   4185
            Left            =   30
            TabIndex        =   9
            Top             =   345
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   7382
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "id"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Apellidos y Nombres"
            Columns(1).DataField=   "apenom"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Usuario"
            Columns(2).DataField=   "login"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Nivel Usuario"
            Columns(3).DataField=   "descripcion"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Correo Electronico"
            Columns(4).DataField=   "email"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   4
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Activo"
            Columns(5).DataField=   "activo"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=503"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=423"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=512"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=5133"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=5054"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2223"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2143"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=2355"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2275"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=3836"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=3757"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=1138"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=1058"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=70,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
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
         Begin VB.Label LblPeriodo 
            Alignment       =   2  'Center
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
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   9810
            TabIndex        =   11
            Top             =   0
            Width           =   1860
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consultando Usuarios"
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
            Left            =   75
            TabIndex        =   10
            Top             =   30
            Width           =   9105
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
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
                  Text            =   "Modificar Usuario"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar Usuario"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Desactivar Usuario"
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
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmManUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo    : FRMMANUSUARIOS
'* Tipo              : FORMULARIO
'* Descripcion       : FORMULARIO PARA EL MANTENIMIENTO DE USUARIOS
'* DISEÑADO POR      : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION   : 03/09/09
'* VERSION           : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstUser As New ADODB.Recordset              ' RECORDSET PRINCIPAL SE USARA PARA CARGAR LA LISTA DE USUARIOS REGOSTRADOS
Dim QueHace As Integer                          ' VARIABLE PARA IDENTIFICAR LAS ACCIONES SOBRE EL FORMULARIO (1 = NUEVO,2 = MODIFICAR, 3 = SOLOLECTURA)
Dim SeEjecuto As Boolean
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO

Dim fOrdenLista As Boolean           ' especfica el orden de la lista de la consulta



Private Sub CmdBusTipuUser_Click()
    ' CARGAMOS LA LISTA PARA BUSCAR EL NIVEL DE USUARIO
    'Dim xform As New EPS_Buscar.Buscar
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_niveluser.* FROM mae_niveluser WHERE id <> 3"
    
    xform.Titulo = "Buscando Nivel de Usuario"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtTipUser.Text = xRs("id")
        LblTipoUser.Caption = xRs("descripcion")
        TxtCorreo.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Dg1_DblClick()
    ' MOSTRAMOS EL TAB NUMERO 2 AL HACER DOBLE CLICK SOBRE EL GRID
    TabOne1.CurrTab = 1
    MuestraSegundoTab
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstUser
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA LA COLUMNAS DEL DtaGrid Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstUser.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstUser("id")), xCon
    End If
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO QUE SE EJECUTARA AL CARGAR EL FORMULARIO, ES AQUI DONDE SE CARGA LA LISTA DE USUARIOS Y SE MUESTRA
    ' EN EL DATA GRID.
    If SeEjecuto = False Then
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = 100
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
                
        RST_Busq RstUser, "SELECT UCase(mae_usuarios.ape)+', '+mae_usuarios.nom AS apenom, mae_niveluser.descripcion, mae_usuarios.*, " _
            & "  mae_niveluser.descripcion FROM mae_usuarios LEFT JOIN mae_niveluser ON mae_usuarios.nivel = mae_niveluser.id ORDER BY " _
            & " UCase(mae_usuarios!ape)+', '+mae_usuarios.nom", xCon
    
        Set Dg1.DataSource = RstUser
        
    End If
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : MuestraSegundoTab()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : MUESTRA LOS DATOS AL DETALLE DEL USUARIO SELECCIONADO
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub MuestraSegundoTab()
    TxtApe.Text = NulosC(RstUser("ape"))
    TxtNom.Text = NulosC(RstUser("nom"))
    TxtTipUser.Text = RstUser("nivel")
    LblTipoUser.Caption = RstUser("descripcion")
    
    TxtCorreo.Text = NulosC(RstUser("email"))
    TxtLogin.Text = NulosC(RstUser("login"))
    TxtContra.Text = NulosC(RstUser("pass"))
    TxtContra2.Text = NulosC(RstUser("passconf"))
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : Bloquea()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : ACTIVA O DESACTIVA LOS CONTROLES DEL FORMULARIO
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Bloquea()
    TxtApe.Locked = Not TxtApe.Locked
    TxtNom.Locked = Not TxtNom.Locked
    TxtTipUser.Locked = Not TxtTipUser.Locked
    TxtCorreo.Locked = Not TxtCorreo.Locked
    TxtLogin.Locked = Not TxtLogin.Locked
    TxtContra.Locked = Not TxtContra.Locked
    TxtContra2.Locked = Not TxtContra2.Locked
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : Blanquea()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : BLANQUEA LOS CONTROLES DEL FORMULARIO PARA EL INGRESO DE DATOS
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Blanquea()
    TxtApe.Text = ""
    TxtNom.Text = ""
    TxtTipUser.Text = ""
    TxtCorreo.Text = ""
    TxtLogin.Text = ""
    TxtContra.Text = ""
    TxtContra2.Text = ""
    LblTipoUser.Caption = ""
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : Nuevo()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : LE DICE AL FORMULARIO QUE SE AGREGARA UN NUEVO REGISTRO, PARA ELLO ACTUALIZA EL
'*                  EL VALOR DE LA VARIABLE QUEHACE = 1
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Nuevo()
    QueHace = 1
    xHorIni = Time
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Blanquea
    Bloquea
    ActivaTool
    Label5.Caption = "Agregando Usuario"
    TxtApe.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : ActivaTool()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : ACTIVA LA BARRA DE HERRAMIENTAS DEL FORMULARIO
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO QUE SE EJECUTARA AL CARGAR EL FORMULARIO
    SeEjecuto = False
    QueHace = 3
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then
            MuestraSegundoTab
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : Nuevo()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : LE DICE AL FORMULARIO QUE SE MODIFICARA UN REGISTRO, PARA ELLO ACTUALIZA EL
'*                  VALOR DE LA VARIABLE QUEHACE = 2
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Modificar()
    QueHace = 2
    xHorIni = Time
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Bloquea
    MuestraSegundoTab
    ActivaTool
    Label5.Caption = "Modificando Datos del Usuario"
    TxtApe.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : Eliminar()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : ELIMINA UN REGISTRO DEL FORMULARIO, PARA ELLO VERIFICA QUE EL USUARIO QUE EFECTUA
'*                  LA ELIMINACION SEA DEL NIVEL 3
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Eliminar()
    If RstUser("nivel") = 3 Then
        MsgBox "No se puede eliminar al usuario master", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Dim Rpta As Integer
    Dim xId As Double
    
    xId = NulosN(RstUser("id"))
    
    Rpta = MsgBox("Esta seguro de desactivar el usuario seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "UPDATE mae_usuarios SET mae_usuarios.activo = 0 WHERE (((mae_usuarios.id)=" & xId & "))"
        
        'grabamos el movimiento en la tabla var_edicion
        GrabarOperacion xIdUsuario, IdMenuActivo, 2, Time, Time, Date, xCon, xId
        
        MsgBox "El usuario se desactivo con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstUser.Requery
        Dg1.Refresh
    End If
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : Grabar()
'* Tipo           : FUNCCION
'* Descripcion    : PERMITE GUDAR LOS DATOS EDITADOS EN EL FORMULARIO, RETORANA UN AVALOR VERDADERO
'*                  CUANDO EL REGISTRO SE GUARDA CON EXITO
'* Paranetros     : NULL
'* Retorna        : BOOLEAN
'*****************************************************************************************************
Function Grabar() As Boolean
    ' VERIFICAMOS QUE LOS DATOS INGRESADO SON LOS CORRECTOS
    If NulosC(TxtApe.Text) = "" Then
        MsgBox "No ha especificado los apellidos del usuario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtApe.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtNom.Text) = "" Then
        MsgBox "No ha especificado los nombres del usuario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNom.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtTipUser.Text) = "" Then
        MsgBox "No ha especificado el tipo de usuario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipUser.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtLogin.Text) = "" Then
        MsgBox "No ha especificado el usuario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtLogin.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtContra.Text) = "" Then
        MsgBox "No ha especificado la contraseña del usuario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtContra.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtContra2.Text) = "" Then
        MsgBox "No ha confirmado la contraseña del usuario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtContra2.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtContra.Text) <> NulosC(TxtContra2.Text) Then
        MsgBox "la contraseña de confirmacion es diferente a la contraseña", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtContra2.SetFocus
        Exit Function
    End If
    
    Dim Rst As New ADODB.Recordset
    Dim xId As Double
    
    ' PREGUNTAMOS QUE OPERACION HA REALIZADO EL FORMULARIO
    If QueHace = 1 Then
        ' SI ES UN NUEVO REGISTRO GENERAMOS UN NUEVO ID
        xId = HallaCodigoTabla("mae_usuarios", xCon, "id")
        RST_Busq Rst, "SELECT * FROM mae_usuarios", xCon
        Rst.AddNew
        Rst("id") = xId
    Else
        ' SI ES UNA ACTUALIZACION DE DATOS, BUSCAMOS EL REGISTRO A MODIFICAR
        xId = RstUser("id")
        RST_Busq Rst, "SELECT * FROM mae_usuarios WHERE id = " & xId & "", xCon
    End If
    
    ' ACTUALIZAMOS LOS DATOS
    Rst("ape") = NulosC(TxtApe.Text)
    Rst("nom") = NulosC(TxtNom.Text)
    Rst("login") = NulosC(TxtLogin.Text)
    Rst("pass") = NulosC(TxtContra.Text)
    Rst("passconf") = NulosC(TxtContra2.Text)
    Rst("email") = NulosC(TxtCorreo.Text)
    Rst("nivel") = NulosN(TxtTipUser.Text)
    If QueHace = 1 Then Rst("activo") = -1
    Rst.Update
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    MsgBox "El usuario se registro con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstUser.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 8 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstUser.Filter = ""
    End If
    
    If Button.Index = 9 Then
        Set RstUser = Nothing
        Unload Me
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then Modificar
    
    If ButtonMenu.Index = 2 Then
        Dim Rpta As Integer
        
        xHorIni = Time
        
        Rpta = MsgBox("Esta seguro de activar el usuario seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
        If Rpta = vbYes Then
            xCon.Execute "UPDATE mae_usuarios SET mae_usuarios.activo = -1 WHERE (((mae_usuarios.id)=" & RstUser("id") & "))"
            
            'grabamos el movimiento en la tabla var_edicion
            GrabarOperacion xIdUsuario, IdMenuActivo, 2, xHorIni, Time, Date, xCon, NulosN(RstUser("id"))
            
            MsgBox "El usuario se activo con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            RstUser.Requery
            Dg1.Refresh
        End If
    End If
End Sub

Private Sub TxtApe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtContra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtContra2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtCorreo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtLogin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtTipUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If NulosC(TxtTipUser.Text) = "" Then
            SendKeys vbTab
            Exit Sub
        End If
        
        LblTipoUser.Caption = Busca_Codigo(NulosN(TxtTipUser.Text), "id", "descripcion", "mae_niveluser", "N", xCon)
        
        If NulosC(LblTipoUser.Caption) = "" Then
            TxtTipUser.Text = ""
        End If
        
        SendKeys vbTab
    End If
End Sub

Private Sub TxtTipUser_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        CmdBusTipuUser_Click
    End If
End Sub

'*****************************************************************************************************
'* Nombre Modulo  : Cancelar()
'* Tipo           : PROCEDIMIENTO
'* Descripcion    : PERMITE CANELAR EL PROCESO DE INGRESO O MODIFICACION DE REGISTRO
'* Paranetros     : NULL
'* Retorna        : NULL
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    Bloquea
    ActivaTool
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
End Sub
