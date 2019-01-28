VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmManCentroCostos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Centro de Costos"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   375
      Width           =   11880
      _cx             =   20955
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6795
         Left            =   45
         TabIndex        =   8
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   30
            TabIndex        =   9
            Top             =   345
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11324
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
            Columns(1).Caption=   "Código"
            Columns(1).DataField=   "codigo"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descripción"
            Columns(2).DataField=   "descripcion"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
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
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2275"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2196"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=9975"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=9895"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
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
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Centro de Costos"
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
            TabIndex        =   10
            Top             =   30
            Width           =   11610
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Detalle de la Cuenta"
         Height          =   6795
         Left            =   12525
         TabIndex        =   1
         Top             =   375
         Width           =   11790
         Begin VB.Frame Frame3 
            Height          =   1905
            Left            =   840
            TabIndex        =   2
            Top             =   2415
            Width           =   10125
            Begin VB.TextBox TxtDescripcion 
               Height          =   300
               Left            =   2460
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   4
               Text            =   "TxtDescripcion"
               Top             =   1005
               Width           =   7050
            End
            Begin VB.TextBox TxtNumCta 
               Height          =   300
               Left            =   2460
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   3
               Text            =   "TxtNumCta"
               Top             =   645
               Width           =   1455
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               Caption         =   "Código"
               Height          =   195
               Left            =   915
               TabIndex        =   6
               Top             =   675
               Width           =   495
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Descripción"
               Height          =   195
               Index           =   1
               Left            =   915
               TabIndex        =   5
               Top             =   1035
               Width           =   840
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle Centro de Costos"
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
            Width           =   11610
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
            Picture         =   "FrmManCentroCostos.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCentroCostos.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCentroCostos.frx":08D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCentroCostos.frx":0C68
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCentroCostos.frx":0DEC
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCentroCostos.frx":1240
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCentroCostos.frx":1358
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCentroCostos.frx":189C
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCentroCostos.frx":1DE0
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCentroCostos.frx":1EF4
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCentroCostos.frx":2008
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCentroCostos.frx":245C
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCentroCostos.frx":25C8
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
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
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   9
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar Excel"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmManCentroCostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstCta As New ADODB.Recordset
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim xHorIni As Date

Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim mIdRegistro& '--identificador del registro
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstCta
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstCta.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstCta("id")), xCon
    End If
End Sub

Private Sub Form_Activate()
'Modificado: 10/01/11 Johan Castro
'            Agregar linea de codigo para bloquear accesos de usuarios
'            Se elimina linea de codigo: CierrePeriodo Toolbar1, 16, 0, False, xCon, xIdUsuario

    If SeEjecuto = False Then
        Dim Rpta As Integer
        
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        '----------------------------------------------
        
        
        RST_Busq RstCta, "SELECT con_centrocosto.* From con_centrocosto where con_centrocosto.id not in (0) " _
            & " ORDER BY con_centrocosto.codigo", xCon
        
        Set Dg1.DataSource = RstCta

    End If
End Sub

Sub Nuevo()
    Bloquea
    Blanquea
    ActivaTool
    QueHace = 1
    Label5.Caption = "Agregando Centro de Costos"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    xHorIni = Time
    TxtNumCta.SetFocus
End Sub

Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
    
'    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
'    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
'    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
'
'    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
'    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
'
'    Toolbar1.Buttons(8).Enabled = Not Toolbar1.Buttons(8).Enabled
'    Toolbar1.Buttons(9).Enabled = Not Toolbar1.Buttons(9).Enabled
'    Toolbar1.Buttons(10).Enabled = Not Toolbar1.Buttons(10).Enabled
'
'    Toolbar1.Buttons(12).Enabled = Not Toolbar1.Buttons(12).Enabled
'    Toolbar1.Buttons(14).Enabled = Not Toolbar1.Buttons(14).Enabled
End Sub

Sub Blanquea()
    TxtNumCta.Text = ""
    TxtDescripcion.Text = ""
End Sub

Sub Bloquea()
    TxtNumCta.Locked = Not TxtNumCta.Locked
    TxtDescripcion.Locked = Not TxtDescripcion.Locked
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    QueHace = 3
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
End Sub

Sub MuestraSegundoTab()
    If RstCta.State = 0 Then Exit Sub
    If RstCta.RecordCount = 0 Then Exit Sub
    TxtNumCta.Text = NulosC(RstCta("codigo"))
    TxtDescripcion.Text = NulosC(RstCta("descripcion"))
End Sub

Sub Cancelar()
    QueHace = 3
    Bloquea
    ActivaTool
    Label5.Caption = "Detalle del Centro de Costo"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Dg1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario mientras este ingresando o modificando un Centro de Costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
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
            RstCta.Requery
            
            TDB_FiltroLimpiar Dg1
            RstCta.Filter = adFilterNone
            
            Dg1.Refresh
            
            RstCta.MoveFirst
            RstCta.Find "id = " & mIdRegistro & ""
            If RstCta.EOF = True Then
                RstCta.MoveFirst
            End If
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 8 Then Filtrar
    
    If Button.Index = 9 Then
        If RstCta.State = 0 Then Exit Sub
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstCta.Filter = adFilterNone
        RstCta.Requery
        Dg1.Refresh
    End If
    
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 12 Then Exportar
    
    If Button.Index = 15 Then
        Set RstCta = Nothing
        Unload Me
    End If
End Sub

Function Grabar() As Boolean
    Dim A As Integer
    
    If TxtNumCta.Text = "" Then
        MsgBox "No ha especificado el codigo del centro de costos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumCta.SetFocus
        Exit Function
    Else
        Dim xLonCad As Integer
        xLonCad = Len(Trim(TxtNumCta.Text))
        
        If xLonCad = 1 Or xLonCad = 3 Or xLonCad = 5 Or xLonCad = 7 Or xLonCad = 9 Then
            MsgBox "El numero de caracteres ingresados no es el correcto para el centro de costos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtNumCta.SetFocus
            Exit Function
        End If
        
        Dim Rst As New ADODB.Recordset
        
        If xLonCad = 4 Then
            RST_Busq Rst, "SELECT con_centrocosto.codigo From con_centrocosto " _
                & " WHERE (((con_centrocosto.codigo)='" & Mid(Trim(TxtNumCta.Text), 1, 2) & "'))", xCon
            If Rst.RecordCount = 0 Then
                MsgBox "El nivel anterior del centro de costos no existe, revise el codigo del centro de costos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                TxtNumCta.SetFocus
                Exit Function
            End If
        End If
        
        If xLonCad = 6 Then
            RST_Busq Rst, "SELECT con_centrocosto.codigo From con_centrocosto " _
                & " WHERE (((con_centrocosto.codigo)='" & Mid(Trim(TxtNumCta.Text), 1, 4) & "'))", xCon
            If Rst.RecordCount = 0 Then
                MsgBox "El nivel anterior del centro de costos no existe, revise el codigo del centro de costos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                TxtNumCta.SetFocus
                Exit Function
            End If
        End If
        
        If xLonCad = 8 Then
            RST_Busq Rst, "SELECT con_centrocosto.codigo From con_centrocosto " _
                & " WHERE (((con_centrocosto.codigo)='" & Mid(Trim(TxtNumCta.Text), 1, 6) & "'))", xCon
            If Rst.RecordCount = 0 Then
                MsgBox "El nivel anterior del centro de costos no existe, revise el codigo del centro de costos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                TxtNumCta.SetFocus
                Exit Function
            End If
        End If
        
        If xLonCad = 10 Then
            RST_Busq Rst, "SELECT con_centrocosto.codigo From con_centrocosto " _
                & " WHERE (((con_centrocosto.codigo)='" & Mid(Trim(TxtNumCta.Text), 1, 8) & "'))", xCon
            If Rst.RecordCount = 0 Then
                MsgBox "El nivel anterior del centro de costos no existe, revise el codigo del centro de costos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                TxtNumCta.SetFocus
                Exit Function
            End If
        End If
        
        If xLonCad = 12 Then
            RST_Busq Rst, "SELECT con_centrocosto.codigo From con_centrocosto " _
                & " WHERE (((con_centrocosto.codigo)='" & Mid(Trim(TxtNumCta.Text), 1, 10) & "'))", xCon
            If Rst.RecordCount = 0 Then
                MsgBox "El nivel anterior del Centro de Costos no existe, revise el codigo del centro de costos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                TxtNumCta.SetFocus
                Exit Function
            End If
        End If
    End If

    If TxtDescripcion.Text = "" Then
        MsgBox "No ha especificado la descripción del Centro de Costos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtDescripcion.SetFocus
        Exit Function
    End If

    Dim RstCab As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    
    Dim xId As Double
        
    On Error GoTo LaCague
    xCon.BeginTrans
    If QueHace = 1 Then
        xId = HallaCodigoTabla("con_centrocosto", xCon, "id")
        RST_Busq RstCab, "SELECT * FROM con_centrocosto", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstCta("id")
        RST_Busq RstCab, "SELECT * FROM con_centrocosto WHERE id = " & RstCta("id") & "", xCon
    End If
    
    mIdRegistro = xId
    
    RstCab("codigo") = TxtNumCta.Text
    RstCab("descripcion") = TxtDescripcion.Text
    
    '*****************************************************************************************************
    '-- 0101
    '-----DEL TIPO   1 = tiene dependencias; 0 = registro
    '--si depende de otro centro de costo
    Dim xRs As New ADODB.Recordset
    Dim nCodigo As String
    nCodigo = Mid(Trim(TxtNumCta.Text), 1, Len(Trim(TxtNumCta.Text)) - 2)
    If nCodigo <> "" Then
        RST_Busq xRs, "SELECT id FROM con_centrocosto WHERE (((codigo)= '" + nCodigo + "'));", xCon
        If xRs.EOF = False Or xRs.BOF = False Or xRs.RecordCount <> 0 Then
            xCon.Execute "UPDATE con_centrocosto SET tipo = 1 WHERE id = " + CStr(xRs.Fields("id")) '--
        End If
    End If
    Set xRs = Nothing
    '--si hay centros de costos que dependen de este
    nCodigo = Trim(TxtNumCta.Text)
    RST_Busq xRs, "SELECT id, codigo FROM con_centrocosto WHERE id <>" + CStr(xId) + " AND codigo Like '" + nCodigo + "%' ;", xCon
    If xRs.EOF = False Or xRs.BOF = False Or xRs.RecordCount <> 0 Then
        RstCab("tipo") = 1
    Else
        RstCab("tipo") = 0
    End If
    '-----
    '*****************************************************************************************************
    RstCab.Update
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId

    MsgBox "El Centro de Costos se guardó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    xCon.CommitTrans
    Grabar = True
    Exit Function
        
LaCague:
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el Centro de Costos por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstCab = Nothing
End Function


Sub Modificar()
    Bloquea
    Blanquea
    ActivaTool
    QueHace = 2
    Label5.Caption = "Modificando Centro de Costos"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    MuestraSegundoTab
    xHorIni = Time
    TxtNumCta.SetFocus
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    Dim xRs As New ADODB.Recordset
    Dim xId As Long
    If RstCta.State = 0 Then Exit Sub
    xId = NulosN(RstCta.Fields("id"))
    If RstCta.RecordCount = 0 Then
        MsgBox "No hay Registros para eliminar", vbExclamation, xTitulo
        Exit Sub
    End If
    
    RST_Busq xRs, "SELECT id, codigo FROM con_centrocosto WHERE (((id)<>" & xId & ") AND ((codigo) Like '" & RstCta.Fields("codigo") & "%'));", xCon
    If xRs.EOF = False Or xRs.BOF = False Or xRs.RecordCount <> 0 Then
        MsgBox "El registro tiene Divisionaria" + vbCr + "Elimine las Divisionarias primero", vbExclamation, xTitulo
        Set xRs = Nothing
        Exit Sub
    End If
    '--verificar si hay centro de costos en compras
    
    
    '--verificar si hay centro de costo en honorario
    
    '--verificar si hay centro de costo en boletas de pago
    
    
    Rpta = MsgBox("Esta seguro de eliminar el Centro de Costo seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        '*****************************************************************************************************
        '-- 0101
        '-----DEL TIPO   1 = tiene dependencias; 0 = registro
        '--si depende de otro centro de costo
        RST_Busq xRs, "SELECT id FROM con_centrocosto WHERE codigo like  '" & Mid(Trim(TxtNumCta.Text), 1, Len(Trim(TxtNumCta.Text)) - 2) & "%';", xCon
        If xRs.RecordCount = 1 Then
            xCon.Execute "UPDATE con_centrocosto SET tipo = 0 WHERE id = " & xId '--
        End If
        Set xRs = Nothing
        '*****************************************************************************************************
        
        xCon.Execute "DELETE * FROM con_centrocosto WHERE id =" & xId & " "
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & xId & " AND idform = " & IdMenuActivo
        
        MsgBox "El Centro de Costo se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstCta.Requery
        Dg1.Refresh
       
        
        
    End If
    
    
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Sub Filtrar()
    Dim xform As New eps_librerias.FormFiltrar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
   
    xCampos(0, 0) = "Código":             xCampos(0, 1) = "codigo":        xCampos(0, 2) = "C":         xCampos(0, 3) = "4200"
    xCampos(1, 0) = "Descripcion":        xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "C":         xCampos(1, 3) = "1500"
    
    Set xform.Coneccion = xCon
    Set xform.Rst = RstCta       'recorset que llena el grid
    Set RstCta = xform.FiltrarReg(xCampos)
    Set Dg1.DataSource = RstCta
    Dg1.Refresh
End Sub

Sub Buscar()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(2, 4) As String
    
    xCampos2(0, 0) = "Numero C.C.":   xCampos2(0, 1) = "codigo":          xCampos2(0, 2) = "1500":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Descripcion":   xCampos2(1, 1) = "descripcion":     xCampos2(1, 2) = "6500":         xCampos2(1, 3) = "C"
    
    xform.SqlCad = "SELECT con_centrocosto.* From con_centrocosto ORDER BY con_centrocosto.descripcion"
    
    xform.Titulo = "Buscando Centro de Costo"
        
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "codigo"
    xform.CampoBusca = "codigo"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos2)
    
    If xRs.State = 1 Then
        RstCta.MoveFirst
        RstCta.Find "id = " & NulosN(xRs("id")) & ""
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub


Sub Exportar()
    Dim oExport As New SGI2_funciones.formularios
    Dim Rst As New ADODB.Recordset
    
    Dim xCampos(2, 3) As String
    
    TabOne1.CurrTab = 0
    '0::Nombre a Mostrar;
    '1::nombre de Campo del Rst;
    '2::alineacion(0::derecha, 1::centro, 2::izquierda);
    '3::ancho de columna
    '--obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Id":                 xCampos(0, 1) = "id":           xCampos(0, 2) = 2:  xCampos(0, 3) = "450"
    xCampos(1, 0) = "Código":             xCampos(1, 1) = "codigo":       xCampos(1, 2) = 0:  xCampos(1, 3) = "1000"
    xCampos(2, 0) = "Descripción":        xCampos(2, 1) = "descripcion":  xCampos(2, 2) = 0:  xCampos(2, 3) = "4500"
        
    Set Rst = RstCta.Clone
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "Centro de Costo", "", "", "Centro de Costos", Rst, xCampos
    Set oExport = Nothing
    Set Rst = Nothing
    Dg1.Refresh
End Sub

