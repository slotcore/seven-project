VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmManConceptoNC_ND 
   Caption         =   "Ventas - Mantenimiento de Conceptos "
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10620
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6750
      Left            =   15
      TabIndex        =   0
      Top             =   405
      Width           =   10605
      _cx             =   18706
      _cy             =   11906
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
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6330
         Left            =   -11160
         TabIndex        =   9
         Top             =   375
         Width           =   10515
         Begin TrueOleDBGrid70.TDBGrid Dg3 
            Height          =   5955
            Left            =   30
            TabIndex        =   10
            Top             =   360
            Width           =   10470
            _ExtentX        =   18468
            _ExtentY        =   10504
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
            Columns(1).Caption=   "Descripción"
            Columns(1).DataField=   "descripcion"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   4
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Hacer Kardex"
            Columns(2).DataField=   "hacekar"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1191"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1111"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=8361"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=8281"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2302"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2223"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
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
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H80&"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
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
         Begin VB.Label LblTipo 
            AutoSize        =   -1  'True
            Caption         =   "LblTipo"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   7650
            TabIndex        =   13
            Top             =   15
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Tipo de Producto"
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
            Left            =   105
            TabIndex        =   11
            Top             =   45
            Width           =   10305
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   6330
         Left            =   45
         TabIndex        =   1
         Top             =   375
         Width           =   10515
         Begin VB.Frame Frame3 
            Height          =   3015
            Left            =   600
            TabIndex        =   2
            Top             =   1470
            Width           =   9375
            Begin VB.OptionButton OptNo 
               Caption         =   "No"
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
               Left            =   2670
               TabIndex        =   15
               Top             =   1665
               Width           =   765
            End
            Begin VB.OptionButton OptSi 
               Caption         =   "Si"
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
               Left            =   1725
               TabIndex        =   14
               Top             =   1665
               Width           =   765
            End
            Begin VB.TextBox txt 
               Height          =   315
               Index           =   1
               Left            =   1710
               MaxLength       =   50
               TabIndex        =   4
               Text            =   "txt(1)"
               Top             =   1215
               Width           =   6330
            End
            Begin VB.TextBox txt 
               BackColor       =   &H0080FF80&
               Height          =   315
               Index           =   0
               Left            =   6810
               TabIndex        =   3
               Tag             =   "null"
               Text            =   "txt(0)"
               Top             =   705
               Visible         =   0   'False
               Width           =   1170
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "Descripción"
               Height          =   195
               Index           =   1
               Left            =   300
               TabIndex        =   7
               Top             =   1260
               Width           =   840
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "Hacer Kardex"
               Height          =   195
               Index           =   2
               Left            =   300
               TabIndex        =   6
               Top             =   1650
               Width           =   975
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "codigo"
               Height          =   195
               Index           =   0
               Left            =   6225
               TabIndex        =   5
               Top             =   825
               Visible         =   0   'False
               Width           =   480
            End
         End
         Begin VB.Label LblTipo2 
            AutoSize        =   -1  'True
            Caption         =   "LblTipo"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   7650
            TabIndex        =   16
            Top             =   15
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Concepto NC"
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
            Left            =   75
            TabIndex        =   8
            Top             =   45
            Width           =   10365
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   10620
      _ExtentX        =   18733
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
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5700
         Top             =   -105
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
               Picture         =   "FrmManConceptoNC_ND.frx":0000
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManConceptoNC_ND.frx":0544
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManConceptoNC_ND.frx":08D6
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManConceptoNC_ND.frx":0A5A
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManConceptoNC_ND.frx":0EAE
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManConceptoNC_ND.frx":0FC6
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManConceptoNC_ND.frx":150A
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManConceptoNC_ND.frx":1A4E
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManConceptoNC_ND.frx":1B62
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManConceptoNC_ND.frx":1C76
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManConceptoNC_ND.frx":20CA
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManConceptoNC_ND.frx":2236
               Key             =   "IMG11"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmManConceptoNC_ND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMMANTIPOPRODUCTO.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : AQUI SE CREAN, MODIFICAN Y ELIMINAN LOS CONCEPTO PARA NOTA DE CREDITO Y DEBITO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 14/12/09
'* VERSION          : 1.0
'*****************************************************************************************************

Option Explicit
Dim QueHace As Integer              ' VARIABLE QUE ESPECIFICA EN QUE ESTADO SE ENCUENTRA EL FORMULARIO 1 = NUEVO; 2 = MODIFICA; 3 = SOLO LECTURA
Dim SeEjecuto As Boolean            ' VARIABLE QUE ESPECIFICA SI SE EJECUTO EL EVENTO ACTIVATE DEL FORMULARIO, SOLO ES USADO EN ESE EVENTO
Dim RstFrm As New ADODB.Recordset   ' RECORDSET UTILIZADO PARA CARGAR TODOS LOS REGISTROS
Dim xTipo As Integer                ' ESPECIFICA EL TIPO DE INFORMACION QUE SE MOSTRARA (1 = NOTA DE CREDITO,  2 = NOTA DE DEBITO)
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO

'Dim Agregando As Boolean

Private Sub Dg3_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    RstFrm.Sort = CStr(Dg3.Columns(ColIndex).DataField)
    Err.Clear
End Sub

Private Sub Dg3_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstFrm("id")), xCon
    End If
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A CARGARSE, ESTE EVENTO SE EJECUTARA UNA SOLA VEZ
    If SeEjecuto = False Then
'        Dim Rpta As Integer
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
    
        CARGAR_GRID
        SeEjecuto = True
'        If RstFrm.RecordCount = 0 Then
'            Rpta = MsgBox("No se ha registrado ningún Tipo de Producto, ¿Desea agergar uno ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
'            If Rpta = vbYes Then
'                Nuevo
'            Else
'                Set RstFrm = Nothing
'                Unload Me
'                Exit Sub
'            End If
'        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : CARGAR_GRID
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGAR EL RECORDSER PRINCIPAL CON DATOS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub CARGAR_GRID()
    Dim xSQL  As String

    If xTipo = 1 Then
        xSQL = "SELECT vta_conceptonc.id, vta_conceptonc.descripcion, vta_conceptonc.hacekar From vta_conceptonc " _
            & " Where (((vta_conceptonc.tipo) = 1)) ORDER BY vta_conceptonc.descripcion"
    Else
        xSQL = "SELECT vta_conceptonc.id, vta_conceptonc.descripcion, vta_conceptonc.hacekar From vta_conceptonc " _
            & " Where (((vta_conceptonc.tipo) = 2)) ORDER BY vta_conceptonc.descripcion"
    End If
         
    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, xSQL, xCon
    
    Set Dg3.DataSource = RstFrm
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    ' SEGUNDO EVENTO A EJECUTARSE
    CentrarFrm Me
    SeEjecuto = False
     QueHace = 3
    Dg3.BatchUpdates = False
    
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    '--
    Habilitar_Obj False
    xTipo = 1
    LblTipo.Caption = "Credito"
    LblTipo2.Caption = "Credito"
    Me.Caption = "Ventas - Mantenimiento de Conceptos de Nota de Credito"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation, xTitulo
        Cancel = 1
        Exit Sub
    End If
    Set Dg3.DataSource = Nothing
End Sub

Private Sub LblTipo_Click()
    If xTipo = 1 Then
        xTipo = 2
        LblTipo.Caption = "Debito"
        LblTipo2.Caption = "Debito"
        Me.Caption = "Ventas - Mantenimiento de Conceptos de Nota de Debito"
    Else
        xTipo = 1
        LblTipo.Caption = "Credito"
        LblTipo2.Caption = "Credito"
        Me.Caption = "Ventas - Mantenimiento de Conceptos de Nota de Credito"
    End If
    CARGAR_GRID
End Sub

Private Sub LblTipo2_Click()
    If xTipo = 1 Then
        xTipo = 2
        LblTipo.Caption = "Debito"
        LblTipo2.Caption = "Debito"
        Me.Caption = "Ventas - Mantenimiento de Conceptos de Nota de Debito"
    Else
        xTipo = 1
        LblTipo.Caption = "Credito"
        LblTipo2.Caption = "Credito"
        Me.Caption = "Ventas - Mantenimiento de Conceptos de Nota de Credito"
    End If
    CARGAR_GRID
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
            If RstFrm.State = 0 Then Exit Sub
            RstFrm.Requery
            Dg3.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 8 Then Filtrar
    If Button.Index = 9 Then RstFrm.Filter = ""
        
    If Button.Index = 11 Then Buscar

    If Button.Index = 15 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA EL REGISTRO ACTUAL
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.BOF = True Or RstFrm.EOF = True Or RstFrm.RecordCount = 0 Then Exit Sub
    
    If xDeDonde = 2 Then Exit Sub '--es unificado
    
    Dim Rpta As Integer
    
    Rpta = MsgBox("¿Esta seguro de eliminar el registro seleccionado?", vbQuestion + vbYesNo, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETe * FROM vta_conceptonc WHERE id = " & RstFrm("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstFrm("id") & " AND idform = " & IdMenuActivo
        
        MsgBox "El registro fue eliminado con éxito", vbInformation, xTitulo
        RstFrm.Requery
        Dg3.Refresh
        If RstFrm.RecordCount = 0 Then
            Rpta = MsgBox("No hay ningún registro, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo)
            If Rpta = vbYes Then
                Nuevo
            Else
                Set RstFrm = Nothing
                Unload Me
                Exit Sub
            End If
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELAR EL PROCESO DE INGRESO O MODIFICACION DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    TabOne1.TabEnabled(0) = True
    ActivaTool
    Habilitar_Obj False
    Label1.Caption = "Detalle del Concepto NC"
    TabOne1.CurrTab = 0
    Dg3.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PERMITE MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Modificar()
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros", vbExclamation, xTitulo
        Exit Sub
    End If

    QueHace = 2
    xHorIni = Time
    TabOne1.TabEnabled(0) = False
    ActivaTool

    Habilitar_Obj True

    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If

    Label1.Caption = "Modificando Concepto NC"
    txt(1).SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LOS DATOS DEL REGISTRO SELECCIONADO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub MuestraSegundoTab()
    On Error GoTo Laca
    With RstFrm
        If .State = 0 Then Exit Sub
        If .EOF = True Or .BOF = True Then Exit Sub
        txt(0).Text = .Fields("id") & ""
        txt(1).Text = .Fields("descripcion") & ""
        'txt(2).Text = .Fields("prefijo") & ""
        If RstFrm("hacekar") = -1 Then
            OptSi.Value = True
        Else
            OptNo.Value = True
        End If
    End With

    Exit Sub
Laca:
    SHOW_ERROR Me.Name, "MuestraSegundoTab"
End Sub

'*****************************************************************************************************
'* Nombre           : Habilitar_Obj
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES TextBox DEL FORMULARIO
'* Paranetros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    band      |  BOOLEAN          |  ESPECIFICA SI SE ACTIVA O DESACTIVA EL CONTROL
'* Devuelve         :
'*****************************************************************************************************
Private Sub Habilitar_Obj(band As Boolean)
    habilitar_Locked txt, Not band
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : BLANQUEA LOS CONTROLES TextBox DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Blanquea()
    LimpiaText txt
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LA BARRA DE HERRAMIENTAS DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub ActivaTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
    
    Toolbar1.Buttons(8).Enabled = Not Toolbar1.Buttons(8).Enabled
    Toolbar1.Buttons(9).Enabled = Not Toolbar1.Buttons(9).Enabled
    Toolbar1.Buttons(10).Enabled = Not Toolbar1.Buttons(10).Enabled
    
    Toolbar1.Buttons(11).Enabled = Not Toolbar1.Buttons(11).Enabled
    
    Toolbar1.Buttons(12).Enabled = Not Toolbar1.Buttons(12).Enabled
    Toolbar1.Buttons(14).Enabled = Not Toolbar1.Buttons(14).Enabled
    Toolbar1.Buttons(15).Enabled = Not Toolbar1.Buttons(15).Enabled
End Sub

'*****************************************************************************************************
'* Nombre           : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PERMITE AGREGAR UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Nuevo()
    QueHace = 1
    xHorIni = Time
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    '---
    Blanquea
    Habilitar_Obj True
    Label1.Caption = "Agregar Concepto NC"
    '-----
    txt(1).SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : FUNCCION
'* Descripcion      : GRABA UN REGISTRO QUE SE ESTA AGREGANDO O MODIFICANCO, DEVUELVE VERDADERO SI
'*                    TIENE EXITO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function Grabar() As Boolean
    ' VALIDAMOS LOS DATOS
    If VALIDAR_DATOS() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " el registro", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo salir
    
    Dim RstCab As New ADODB.Recordset
    Dim xCod As Double
    Dim xCol, xFil As Integer
    
    On Error GoTo LaCague

    xCon.BeginTrans
    
    If QueHace = 1 Then
        ' SI ES UN NUEVO REGISTRO OBTENEMOS EL ULTIMO ID E INCRMENTAMOS U VALOR EN UN
        xCod = HallaCodigoTabla("vta_conceptonc", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM vta_conceptonc ", xCon
        RstCab.AddNew
        RstCab("id") = xCod
        txt(0).Text = xCod
    Else
        ' SI ES UNA ACTUALIACION BUSCAMOS EL ID DEL REGISTRO Y ACTUALIZAMOS LOS DATOS
        RST_Busq RstCab, "SELECT * FROM vta_conceptonc WHERE id =" & RstFrm("id") & "", xCon
        xCod = RstFrm("id")
    End If
    
    RstCab("descripcion") = Trim(txt(1).Text)
    RstCab("tipo") = xTipo
    If OptSi.Value = True Then
        RstCab("hacekar") = -1
    Else
        RstCab("hacekar") = 0
    End If

    RstCab.Update
'    '*************************************************************************************
'    '*** SINCRONIZAR BASE DE DATOS - vta_conceptonc ***'
'    If xDeDonde = 2 Then SincronizarBD xCon, "vta_conceptonc", xCod, QueHace
'    '*************************************************************************************

    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xCod


    MsgBox "El concepto de nota de " & LblTipo.Caption & " se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    xCon.CommitTrans
    Grabar = True
    Exit Function
    
salir:
    Set RstCab = Nothing
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    SHOW_ERROR Me.Name, "Grabar", True
    Exit Function
End Function

'*****************************************************************************************************
'* Nombre           : VALIDAR_DATOS
'* Tipo             : FUNCION
'* Descripcion      : VALIDA QUE EL NUEVO REGISTRO NO EXISTA EN LA TABLA vta_conceptonc, DEVUELVE
'*                    VERDADERO SI ENCUENTRA EL REGISTRO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Private Function VALIDAR_DATOS() As Boolean
    '--SI EL REGISTRO YA EXISTE
    Dim RstTmp As New ADODB.Recordset
    If QueHace = 1 Then
        RST_Busq RstTmp, "SELECT vta_conceptonc.id From vta_conceptonc WHERE ((ucase(trim(vta_conceptonc.descripcion))='" + UCase(Trim(txt(1).Text)) + "'));", xCon
    Else
        RST_Busq RstTmp, "SELECT vta_conceptonc.id FROM vta_conceptonc WHERE (((vta_conceptonc.id)<>" + CStr(RstFrm.Fields("id")) + ") AND (ucase(trim(vta_conceptonc.descripcion))='" + UCase(Trim(txt(1).Text)) + "'));", xCon
    End If
    If RstTmp.BOF = False Or RstTmp.EOF = False Or RstTmp.RecordCount <> 0 Then
        MsgBox "El registro " + IIf(QueHace = 1, " ya fue ingresado", "ya existe"), vbExclamation, xTitulo
        Set RstTmp = Nothing
        Exit Function
    End If
    Set RstTmp = Nothing
        
    VALIDAR_DATOS = True
End Function
 
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then Imprimir True
    If ButtonMenu.Index = 2 Then Imprimir
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : BUSCA UN REGISTRO EN EL RECORSET PRINCIPAL RstFrm
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Buscar()
    On Error GoTo error
    TabOne1.CurrTab = 0
     
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
   
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripción":     xCampos(0, 1) = "Descripcion":   xCampos(0, 2) = "3500":     xCampos(0, 3) = "C"
    xCampos(1, 0) = "codigo":          xCampos(1, 1) = "id":            xCampos(1, 2) = "1000":     xCampos(1, 3) = "C"
    
    If xTipo = 1 Then
        nSQL = "SELECT vta_conceptonc.id, vta_conceptonc.descripcion FROM vta_conceptonc WHERE tipo = 1 ORDER BY vta_conceptonc.descripcion"
    Else
        nSQL = "SELECT vta_conceptonc.id, vta_conceptonc.descripcion FROM vta_conceptonc WHERE tipo = 2 ORDER BY vta_conceptonc.descripcion"
    End If
         
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Concepto", "Descripcion", "Descripcion", Principio
    If xRs.State = 0 Then GoTo salir
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo salir
    
    RstFrm.MoveFirst
    RstFrm.Find "id = " + CStr(xRs("id"))
    Exit Sub
    
salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub

'*****************************************************************************************************
'* Nombre           : Filtrar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : FILTRA EL RECORSER PRINCIPAL RstFrm, PARA REALIZAR UNA BUSQUEDA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Filtrar()
    Dim xCampos(1, 4) As String
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "Descripcion":   xCampos(0, 2) = "C":     xCampos(0, 3) = "1500"
    xCampos(1, 0) = "Codigo":           xCampos(1, 1) = "id":            xCampos(1, 2) = "N":     xCampos(1, 3) = "800"
    
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg3
    TabOne1.CurrTab = 0
End Sub

'*****************************************************************************************************
'* Nombre           : Imprimir
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPRIME LOS REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Imprimir(Optional IMP_LISTADO As Boolean = False)
    On Error GoTo error
    Me.MousePointer = vbHourglass
    If IMP_LISTADO = False Then
        If Me.TabOne1.CurrTab = 0 Then
        Else
        End If
    Else
        TDB_IMPRIMIR Dg3, "IMPRESIÓN DE CONCEPTO NOTA DE " & UCase(LblTipo.Caption), ""
    End If
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "IMPRIMIR"
End Sub

