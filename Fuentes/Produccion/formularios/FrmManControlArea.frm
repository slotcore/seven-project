VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManControlArea 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Producción - Distribución de Tareas por Areas"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   11820
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6570
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
            Picture         =   "FrmManControlArea.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlArea.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlArea.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlArea.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlArea.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlArea.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlArea.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlArea.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlArea.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlArea.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlArea.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlArea.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManControlArea.frx":277E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
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
            Object.Visible         =   0   'False
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
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Listado"
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
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7065
      Left            =   15
      TabIndex        =   1
      Top             =   375
      Width           =   11790
      _cx             =   20796
      _cy             =   12462
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
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6645
         Left            =   12435
         TabIndex        =   5
         Top             =   375
         Width           =   11700
         Begin VB.Frame Frame9 
            Height          =   1725
            Left            =   45
            TabIndex        =   8
            Top             =   300
            Width           =   11610
            Begin VB.CommandButton cmd 
               Caption         =   "Buscar Tarea"
               Enabled         =   0   'False
               Height          =   465
               Index           =   4
               Left            =   10050
               TabIndex        =   14
               ToolTipText     =   "Buscar Tarea"
               Top             =   1200
               Width           =   1410
            End
            Begin VB.CommandButton cmd 
               Caption         =   "Eliminar Tarea"
               Enabled         =   0   'False
               Height          =   465
               Index           =   3
               Left            =   10050
               TabIndex        =   13
               ToolTipText     =   "Eliminar Tarea"
               Top             =   675
               Width           =   1410
            End
            Begin VB.CommandButton cmd 
               Caption         =   "Agregar Tareas Faltantes"
               Enabled         =   0   'False
               Height          =   465
               Index           =   2
               Left            =   10050
               TabIndex        =   12
               ToolTipText     =   "Agregar Tarea Faltantes"
               Top             =   150
               Width           =   1410
            End
            Begin VB.CommandButton cmd 
               Caption         =   "Eliminar Area"
               Enabled         =   0   'False
               Height          =   450
               Index           =   1
               Left            =   8445
               TabIndex        =   10
               TabStop         =   0   'False
               ToolTipText     =   "Eliminar Tarea"
               Top             =   1080
               Width           =   1410
            End
            Begin VB.CommandButton cmd 
               Caption         =   "Agregar Area"
               Enabled         =   0   'False
               Height          =   465
               Index           =   0
               Left            =   8445
               TabIndex        =   9
               ToolTipText     =   "Agregar Tarea"
               Top             =   510
               Width           =   1410
            End
            Begin VSFlex7Ctl.VSFlexGrid Fg1 
               Height          =   1515
               Left            =   60
               TabIndex        =   11
               Top             =   150
               Width           =   8250
               _cx             =   14552
               _cy             =   2672
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
               Rows            =   20
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManControlArea.frx":2B10
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
            Begin VB.Line Line2 
               BorderColor     =   &H00FFFFFF&
               BorderWidth     =   2
               Index           =   1
               X1              =   9945
               X2              =   9945
               Y1              =   120
               Y2              =   1575
            End
            Begin VB.Line Line2 
               Index           =   0
               X1              =   9915
               X2              =   9915
               Y1              =   150
               Y2              =   1605
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg2 
            Height          =   4575
            Left            =   45
            TabIndex        =   7
            Top             =   2070
            Width           =   11640
            _cx             =   20532
            _cy             =   8070
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
            ForeColorSel    =   16777215
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
            Rows            =   2
            Cols            =   8
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmManControlArea.frx":2BDF
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
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Distribución de Tareas por Area"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   75
            TabIndex        =   6
            Top             =   75
            Width           =   11475
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6645
         Left            =   45
         TabIndex        =   2
         Top             =   375
         Width           =   11700
         Begin TrueOleDBGrid70.TDBGrid Dg3 
            Height          =   6225
            Left            =   30
            TabIndex        =   3
            Top             =   345
            Width           =   11565
            _ExtentX        =   20399
            _ExtentY        =   10980
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
            Columns(1).Caption=   "IdArea"
            Columns(1).DataField=   "idarea"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Area"
            Columns(2).DataField=   "area"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Abrev"
            Columns(3).DataField=   "abrev"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Encargado de Area"
            Columns(4).DataField=   "nombres"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Nº Tareas"
            Columns(5).DataField=   "tottarea"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=794"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=714"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1111"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1032"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Visible=0"
            Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(15)=   "Column(2).Width=4842"
            Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=4763"
            Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(21)=   "Column(3).Width=1614"
            Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=1535"
            Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(27)=   "Column(4).Width=7461"
            Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._WidthInPix=7382"
            Splits(0)._ColumnProps(30)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(31)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(32)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(33)=   "Column(5).Width=1958"
            Splits(0)._ColumnProps(34)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._WidthInPix=1879"
            Splits(0)._ColumnProps(36)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(37)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(38)=   "Column(5).Order=6"
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
            HeadLines       =   1.25
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
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
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
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Distribución de Tareas por Area"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   4
            Top             =   45
            Width           =   11520
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
      End
   End
End
Attribute VB_Name = "FrmManControlArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMMANCONTROLAREA.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : PERMITE ASIGNAR TAREAS A UN AREA
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 02/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim QueHace As Integer              ' INDICA EL ESTADO ACTUAL DEL FORMULARIO
Dim RstFrm As New ADODB.Recordset   ' RECORDSET QUE ALMACENARA LOS DATOS DE LAS TAREAS
Dim Mostrando As Boolean            ' ESPECIFICA QUE SE ESTA AGREGANDO FILAS AL CONTROL FLEXGRID
Dim SeEjecuto As Boolean            ' VERIFICA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim Agregando As Boolean
Dim mIdRegistro&                    ' identificador del registro
Dim fOrdenLista As Boolean          ' especfica el orden de la lista de la consulta

'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA LA BUSQUEDA DE UNA PERSONA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Buscar()
    On Error GoTo error
    TabOne1.CurrTab = 0
     
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Area":         xCampos(0, 1) = "area":     xCampos(0, 2) = "3500":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Encargado":    xCampos(1, 1) = "nombres":  xCampos(1, 2) = "2500":    xCampos(1, 3) = "C"
    
    nSQL = "SELECT pro_area.id, mae_area.descripcion AS area, mae_area.abrev, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, pro_area.idarea, pro_area.idper, (select count(pro_areadet.idar) from pro_areadet where pro_areadet.idar = pro_area.id and pro_areadet.activo=-1) AS tottarea " _
        + vbCr + " FROM (pla_empleados RIGHT JOIN (pro_area LEFT JOIN pro_emp ON pro_area.idper = pro_emp.id) ON pla_empleados.id = pro_emp.idemp) LEFT JOIN mae_area ON pro_area.idarea = mae_area.id;"
        
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscanco Personal", "nomemp", "nomemp", Principio
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    RstFrm.MoveFirst
    RstFrm.Find "id = " + CStr(xRs("id"))

SALIR:
    Set xRs = Nothing
    Exit Sub

error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO EN LA PESTAÑA DETALLE DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    If RstFrm.RecordCount = 0 Then Exit Sub
    Agregando = True
    Dim nSQL As String
    Dim RstTmp As New ADODB.Recordset
    
    Me.MousePointer = vbHourglass
    Fg1.Rows = 1
    With RstFrm
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(.Fields("area"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(.Fields("abrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(.Fields("nombres"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosN(.Fields("idarea"))
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosN(.Fields("idper"))
            .MoveNext
        Loop
    End With
    
    ' reiniciar la cantidad de filas del grid2
    Fg2.Rows = Fg2.FixedRows
    
    If Fg1.Rows > 1 Then
        Set RstTmp = Nothing
        Dim nSQLPivot As String
        Dim mCol&, mFil&
        
        ' generar los id's de las areas para generar la consulta segun orden
        nSQLPivot = GRID_GENERAR_SQL_ID(Fg1, 4, " PIVOT pro_area.idarea ", " IN ", True)
        nSQL = "TRANSFORM Sum(pro_areadet.activo) AS SumaDeactivo " _
            + vbCr + " SELECT pro_tareas.descripcion AS tarea, mae_unidades.abrev, IIf([pro_tareas].[diverso]=0,'No','Si') AS esdiverso, pro_tareas.id " _
            + vbCr + " FROM mae_unidades RIGHT JOIN (((pro_area INNER JOIN pro_areadet ON pro_area.id = pro_areadet.idar) LEFT JOIN mae_area ON pro_area.id = mae_area.id) LEFT JOIN pro_tareas ON pro_areadet.idtar = pro_tareas.id) ON mae_unidades.id = pro_tareas.idunimed " _
            + vbCr + " WHERE pro_tareas.id IS NOT NULL " _
            + vbCr + " GROUP BY pro_tareas.descripcion, mae_unidades.abrev, IIf([pro_tareas].[diverso]=0,'No','Si'), pro_tareas.id, pro_tareas.diverso, pro_tareas.descripcion " _
            + vbCr + " ORDER BY pro_tareas.diverso DESC , pro_tareas.descripcion " _
            + vbCr + nSQLPivot
           
        RST_Busq RstTmp, nSQL, xCon
        ' configurando la cantidad de columnas
        Fg2.Cols = RstTmp.Fields.Count + 1
        
        ' colocando los encabezados de las areas
        For mFil = 1 To Fg1.Rows - 1
            Fg2.TextMatrix(0, mFil + 4) = Fg1.TextMatrix(mFil, 2)      ' abrev
            Fg2.TextMatrix(1, mFil + 4) = Fg1.TextMatrix(mFil, 4)      ' idarea
        Next mFil
        
        ' poner formato del grid2
        For mCol = 5 To Fg2.Cols - 1
            Fg2.ColDataType(mCol) = flexDTBoolean
            Fg2.ColAlignment(mCol) = flexAlignCenterCenter
            Fg2.Row = 0: Fg2.Col = mCol: Fg2.CellAlignment = flexAlignCenterCenter
            Fg2.ColWidth(mCol) = 700
        Next mCol
        
        With RstTmp
            If .RecordCount > 0 Then .MoveFirst
            Do While Not .EOF
                Fg2.Rows = Fg2.Rows + 1
                
                ' recorriendo todas las columnas
                For mCol = 0 To RstTmp.Fields.Count - 1
                    Fg2.TextMatrix(Fg2.Rows - 1, mCol + 1) = NulosC(.Fields(mCol))
                Next mCol
                .MoveNext
            Loop
        End With
        Set RstTmp = Nothing
    End If
    
    GRID_AGRUPAR Fg1, 4     ' poner los colores en la lista de areas
    GRID_AGRUPAR Fg2, 4     ' poner los colores en la lista de taras

    Agregando = False
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Agregando = False
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "MuestraSegundoTab"
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS BOTONES DE LA BARRA DE HERRAMIENTAS DEL FORMULARIO
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
'* Nombre           : pHabilitarObj
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES ESPECIFICADOS
'* Paranetros       : NOMBRE    |  TIPO     |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    band      |  Boolean  |  ESPECIFICA SE SE ACTIVA O DESACTIVA EL CONTROL
'* Devuelve         :
'*****************************************************************************************************
Private Sub pHabilitarObj(band As Boolean)
    habilitar Cmd, band
    Cmd(4).Enabled = True
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : BLANQUEA LAS FILAS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    Fg1.Rows = 1
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL PROCESO DE INGRESO O MODIFICACION DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    pHabilitarObj False
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    ActivaTool
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA pro_area, ESTA FUNCION DEVUELVE VERDADERO CUANDO
'*                    TIENE EXITO
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Function Grabar() As Boolean
    If fValidarDatos() = False Then Exit Function
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " la distribución de las tareas", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xId, xFil&, xFil1&
    
    On Error GoTo LaCague
    xCon.BeginTrans
    Me.MousePointer = vbHourglass
    
    ' eliminamos todos los registros
    xCon.Execute "Delete From pro_areadet"
    xCon.Execute "Delete From pro_area"
    
    ' cargamos el rst
    RST_Busq RstCab, "SELECT TOP 1 * FROM pro_area", xCon
    RST_Busq RstDet, "SELECT TOP 1 * FROM pro_areadet", xCon
    
    For xFil = 1 To Fg1.Rows - 1
        RstCab.AddNew
        RstCab("id") = xFil
        RstCab("idarea") = NulosN(Fg1.TextMatrix(xFil, 4))
        RstCab("idper") = NulosN(Fg1.TextMatrix(xFil, 5))
        RstCab.Update
        ' xFil + 4:: indica la posicion del area
        For xFil1 = 2 To Fg2.Rows - 1
            RstDet.AddNew
            RstDet("idar") = xFil
            RstDet("idtar") = NulosN(Fg2.TextMatrix(xFil1, 4))
            RstDet("activo") = NulosN(Fg2.TextMatrix(xFil1, xFil + 4))
            RstDet.Update
        Next xFil1
    Next xFil
    Me.MousePointer = vbDefault
    xCon.CommitTrans
    Grabar = True
    MsgBox "La distribución de las tareas se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    Set RstCab = Nothing:  Set RstDet = Nothing
    Exit Function

LaCague:
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:  Set RstDet = Nothing
    xCon.RollbackTrans
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo "
End Function

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA pro_area
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    On Error GoTo error
    Dim Rpta As Integer
    Dim xId&
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        If RstFrm.RecordCount = 0 Then
            MsgBox "No hay registros", vbExclamation, xTitulo
        Else
            MsgBox "Seleccione un Registro para Eliminar", vbExclamation, xTitulo
        End If
        Exit Sub
    End If
    xId = NulosN(RstFrm("id"))
    Rpta = MsgBox("¿Esta seguro de eliminar el registro seleccionado?", vbQuestion + vbYesNo, xTitulo)
    
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM pro_areadet WHERE idar = " & xId & ""
        xCon.Execute "DELETE * FROM pro_area WHERE id = " & xId & ""
        RstFrm.Requery
        Dg3.Refresh
        MsgBox "Registro fue eliminado con éxito", vbInformation + vbOKOnly, xTitulo
    End If
    
    TabOne1.CurrTab = 0
    Exit Sub

error:
    SHOW_ERROR Me.Name, "Eliminar", True, "Error al eliminar..."
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA LA EDICION DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    TabOne1.CurrTab = 1
    ActivaTool
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Modificando la Distribución de Tareas por Area"
    QueHace = 2
    pHabilitarObj True
    Blanquea
    MuestraSegundoTab
End Sub

'*****************************************************************************************************
'* Nombre           : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Nuevo()
    QueHace = 1
    TabOne1.CurrTab = 1
    ActivaTool
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Agregando la Distribución de Tareas por Area"
    pHabilitarObj True
    Blanquea
    pConfigurarGrilla
    
    ' cargar la lista de tareas por defecto
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    
    nSQL = "SELECT pro_tareas.descripcion, IIf([pro_tareas].[diverso]=0,'No','Si') AS esdiverso, mae_unidades.abrev, pro_tareas.id " _
        + vbCr + " FROM mae_unidades RIGHT JOIN pro_tareas ON mae_unidades.id = pro_tareas.idunimed " _
        + vbCr + " ORDER BY pro_tareas.diverso DESC , pro_tareas.descripcion; "
    
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    Agregando = True
    Fg2.Rows = Fg2.FixedRows
    Do While Not RstTmp.EOF
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(RstTmp("descripcion"))
        Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosC(RstTmp("abrev"))
        Fg2.TextMatrix(Fg2.Rows - 1, 3) = NulosC(RstTmp("esdiverso"))
        Fg2.TextMatrix(Fg2.Rows - 1, 4) = NulosN(RstTmp("id"))
        RstTmp.MoveNext
    Loop
    Set RstTmp = Nothing
    Agregando = False
    Fg1.SelectionMode = flexSelectionFree
    Fg2.SelectionMode = flexSelectionFree
    Cmd(0).SetFocus
End Sub

Private Sub Dg3_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg3_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA EN FORMA ASCENDENTE O DECENDENTE LAS COLUMNAS DE UN DATA GRID
On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg3.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If QueHace = 3 Then Exit Sub
    If Agregando = True Then Exit Sub
    If Col <> 4 Then Exit Sub
    
    ' fg2.col = fg1.row +4
    ' comparar si es momento de agregar una columna adicional
    If Fg2.Cols <= Fg1.Row + 4 Then
        Fg2.Cols = Fg2.Cols + 1
        Fg2.ColDataType(Fg2.Cols - 1) = flexDTBoolean
        Fg2.ColAlignment(Fg2.Cols - 1) = flexAlignCenterCenter
        Fg2.Row = 0: Fg2.Col = Fg2.Cols - 1: Fg2.CellAlignment = flexAlignCenterCenter
        Fg2.ColWidth(Fg2.Cols - 1) = 700
    End If
    Fg2.TextMatrix(0, Fg1.Row + 4) = Fg1.TextMatrix(Row, 2)
    Fg2.TextMatrix(1, Fg1.Row + 4) = Fg1.TextMatrix(Row, 4)
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    If Fg1.Col = 1 Or Fg1.Col = 3 Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col <> 1 And Col <> 3 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLNotId As String
    Dim nTitulo As String

    Select Case Col
        Case 1 ' area
            ReDim xCampos(3, 4) As String
            xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":   xCampos(0, 2) = "3500": xCampos(0, 3) = "C":
            xCampos(1, 0) = "Abrev":        xCampos(1, 1) = "abrev":    xCampos(1, 2) = "1000":  xCampos(1, 3) = "C":
            xCampos(2, 0) = "Id":           xCampos(2, 1) = "id":       xCampos(2, 2) = "600":  xCampos(2, 3) = "N":
            ' generar la consulta de las areas que no se consideraran, para no generar duplicados
            nSQLNotId = GRID_GENERAR_SQL_ID(Fg1, 4, " WHERE mae_area.id ", "NOT IN", True)
            ' armar la consulta
            nSQL = "SELECT mae_area.id, mae_area.descripcion as nombre,mae_area.abrev  " _
                + vbCr + " FROM mae_area " & nSQLNotId
            nTitulo = "Buscando Area"
                            
        Case 3 ' encargado
            If NulosN(Fg1.TextMatrix(Row, 4)) = 0 Then '--
                MsgBox "Seleccione el Area", vbExclamation, xTitulo
                Fg1.Col = 1
                Fg1.SetFocus
                Exit Sub
            End If
            
            ReDim xCampos(1, 4) As String
            xCampos(0, 0) = "Apellidos y Nombres":  xCampos(0, 1) = "nombre":      xCampos(0, 2) = "4500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            
            nSQL = "SELECT pro_emp.idemp, pro_emp.id AS idper, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, mae_dociden.abrev, pla_empleados.numdoc " _
                + vbCr + " FROM mae_dociden RIGHT JOIN ((pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) ON mae_dociden.id = pla_empleados.idtipdoc " _
                + vbCr + " WHERE (((pro_empdet.idfun)=5));"
            ' fun=5::encargado de area (ver tabla pro_funcion)
            nTitulo = "Buscando Encargados de Area"
    End Select
    
    ' mostrar el formulario buscar
    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio, ""
    
    If RstTmp.State = 0 Then GoTo SALIR
    If RstTmp.EOF = True Or RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR
    Agregando = True
    
    If Col = 1 Then         ' area
        Fg1.TextMatrix(Row, 1) = NulosC(RstTmp.Fields("nombre"))
        Fg1.TextMatrix(Row, 2) = NulosC(RstTmp.Fields("abrev"))
        Agregando = False
        Fg1.TextMatrix(Row, 4) = NulosN(RstTmp.Fields("id"))
        Agregando = True
        GRID_AGRUPAR Fg2, 4 ' poner los colores en la lista de taras
    ElseIf Col = 3 Then     ' encargado
        Fg1.TextMatrix(Row, 3) = NulosC(RstTmp.Fields("nombre"))
        Fg1.TextMatrix(Row, 5) = NulosN(RstTmp.Fields("idper"))
    End If
    Agregando = False
    Set RstTmp = Nothing
    Exit Sub

SALIR:
    Set RstTmp = Nothing
    Agregando = False
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii <> 13 Then KeyAscii = 0
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo error
    If KeyCode = 114 Or KeyCode = vbKeyInsert Then   ' F3 = Agregar Item
        cmd_Click 0
    End If
    
    If KeyCode = 115 Or KeyCode = vbKeyDelete Then   ' F4 = Eliminar Item
        cmd_Click 1
    End If
    Exit Sub

error:
    SHOW_ERROR Me.Name, "Fg1_KeyUp"
End Sub

Private Sub Fg2_EnterCell()
    If QueHace = 3 Then
        Fg2.Editable = flexEDNone
        Exit Sub
    End If
    
    If Fg2.Col >= 5 Then
        Fg2.Editable = flexEDKbdMouse
    Else
        Fg2.Editable = flexEDNone
    End If
End Sub

Private Sub Fg2_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo error
    If KeyCode = 115 Or KeyCode = vbKeyDelete Then   ' F4 = Eliminar tarea
        cmd_Click 3
    End If
    Exit Sub

error:
    SHOW_ERROR Me.Name, "Fg2_KeyUp"
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        pConfigurarGrilla
        Cmd(4).Enabled = True ' habilitar las la opcion de buscar tarea
        Dim nSQL As String
        nSQL = "SELECT pro_area.id, mae_area.descripcion AS area, mae_area.abrev, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, pro_area.idarea, pro_area.idper, (select count(pro_areadet.idar) from pro_areadet where pro_areadet.idar = pro_area.id and pro_areadet.activo=-1) AS tottarea " _
            + vbCr + " FROM (pla_empleados RIGHT JOIN (pro_area LEFT JOIN pro_emp ON pro_area.idper = pro_emp.id) ON pla_empleados.id = pro_emp.idemp) LEFT JOIN mae_area ON pro_area.idarea = mae_area.id;"

        RST_Busq RstFrm, nSQL, xCon
        
        Set Dg3.DataSource = RstFrm
        If RstFrm.RecordCount = 0 Then
            Dim Rpta As Integer
            Rpta = MsgBox("La Distribución de las tareas está vacía, ¿Desea Configurarlo ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                Modificar
            Else
                Set RstFrm = Nothing
                Unload Me
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    QueHace = 3
    SeEjecuto = False
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Mostrando = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RstFrm = Nothing
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
            RstFrm.Requery
            Dg3.Refresh
            If RstFrm.RecordCount <> 0 Then
                RstFrm.MoveFirst
                RstFrm.Find "id=" & mIdRegistro
                If RstFrm.EOF = True Then RstFrm.MoveFirst
            End If
            Cancelar
        End If
    End If
    
    If Button.Index = 6 Then
        Cancelar
    End If
    
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 12 Then pExportar
    
    If Button.Index = 13 Then TDB_IMPRIMIR Dg3, "IMPRESIÓN", "LISTADO DE AREAS DE PRODUCCIÓN"
        
    If Button.Index = 15 Then
        Unload Me
        Set RstFrm = Nothing
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : fValidarDatos
'* Tipo             : FUNCION
'* Descripcion      : VALIDA QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Private Function fValidarDatos() As Boolean
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado el registro de las areas", vbInformation, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    
    ' validar la grilla
    Dim mRow&, mCol&
    
    mCol = -1
    For mRow = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(mRow, 4)) = 0 Then                    ' area
            MsgBox "Falta seleccionar el Area", vbExclamation, xTitulo
            mCol = 1
        End If
        If mCol <> -1 Then Exit For
    Next mRow
    
    If mCol <> -1 Then
        Agregando = True:  Fg1.Row = mRow: Fg1.Col = mCol: Agregando = False
        Fg1.SetFocus
        Exit Function
    End If
    fValidarDatos = True
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
    
    Dim nSQL As String
    Dim oExport As New SGI2_funciones.formularios
    Dim RstTmp  As New ADODB.Recordset
    Dim xCampos(3, 3) As String
    
    ' 0 = Nombre a Mostrar;
    ' 1 = nombre de Campo del Rst;
    ' 2 = alineacion(0::derecha, 1::centro, 2::izquierda);
    ' 3 = ancho de columna
    ' obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Area":       xCampos(0, 1) = "area":    xCampos(0, 2) = 0:  xCampos(0, 3) = "2000"
    xCampos(1, 0) = "Encargado":  xCampos(1, 1) = "nombres": xCampos(1, 2) = 0:  xCampos(1, 3) = "3500"
    Set RstTmp = RstFrm.Clone
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "Lista de Areas de Producción", "", "", "Areas de Producción", RstTmp, xCampos
    Set oExport = Nothing
    Set RstTmp = Nothing
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0 ' agregar
            pRegistroAdd
            
        Case 1 ' eliminar
            pRegistroDel True
            
        Case 2 ' cargar tareas faltantes
            pCargarTareasFaltantes
            
        Case 3 ' eliminar tarea
            pRegistroDel False
            
        Case 4 ' buscar tarea
            pBuscarVSFlexGrid
    End Select
End Sub

'*****************************************************************************************************
'* Nombre           : pRegistroAdd
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : AGREGA UNA FILA AL CONTROL FLEXGRID Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pRegistroAdd()
    Dim mCol%
    Dim fInsertar As Boolean
    If QueHace = 3 Then Exit Sub
    Agregando = True
    If Fg1.Rows > Fg1.FixedRows Then
        If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 4)) = 0 Then      ' tipo de persona
            MsgBox "Seleccione el Area", vbExclamation, xTitulo
            mCol = 1
        Else
            fInsertar = True
            mCol = 1
        End If
    Else
        fInsertar = True
        mCol = 1
    End If
    
    If fInsertar = True Then Fg1.AddItem ""
    
    Fg1.Row = Fg1.Rows - 1
    Fg1.Col = mCol
    
    ' cargar el buscador por defecto
    If fInsertar = True Then Fg1_CellButtonClick Fg1.Rows - 1, 1

    Fg1.SetFocus
    Agregando = False
End Sub

'*****************************************************************************************************
'* Nombre           : pRegistroDel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UNA FILA DEL CONTROL FLEXGRID Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pRegistroDel(Optional fDesdeArea As Boolean = True)
    If QueHace = 3 Then Exit Sub
    If fDesdeArea = True Then
        If Fg1.Row < 1 Then
            MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        ElseIf Fg1.Rows = 1 Then
            MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Fg1.SetFocus
            Exit Sub
        End If
        If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
        
        ' primero eliminamos la columna de las areas vs tareas
        GRID_DELETE Fg2, Fg1.Row + 4, Fg1.Row + 4, e_Columna
        
        ' eliminando el registro
        Fg1.RemoveItem Fg1.Row
        If Fg1.Rows > 1 Then
            Fg1.Row = Fg1.Rows - 1
            Fg1.Col = 1
            Fg1.SetFocus
        Else
            Cmd(0).SetFocus
        End If
    Else
        If Fg2.Row < 1 Then
            MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        ElseIf Fg2.Rows = 1 Then
            MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Fg1.SetFocus
            Exit Sub
        End If
        
        If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
        ' eliminando el registro
        Fg2.RemoveItem Fg2.Row
        
        If Fg2.Rows > 1 Then
            Fg2.Row = Fg2.Rows - 1
            Fg2.Col = 1
            Fg2.SetFocus
        Else
            Cmd(2).SetFocus
        End If
    End If
    
End Sub

'*****************************************************************************************************
'* Nombre           : pConfigurarGrilla
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CONFIGURA LA CABECERA DEL CONTROL FLEXGRID Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pConfigurarGrilla()
    Dim RstTmp As New ADODB.Recordset
    
    With Fg1 ' de las Areas
        .Rows = 1
        .Cols = 6
        .FixedRows = 1
        .RowHeight(0) = 300
        .TextMatrix(0, 1) = "Area":      .ColWidth(1) = 2500:  .ColAlignment(1) = flexAlignLeftCenter:  .Row = 0: .Col = 1: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 2) = "Abrev":     .ColWidth(2) = 750:   .ColAlignment(2) = flexAlignLeftCenter:  .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 3) = "Encargado": .ColWidth(3) = 4500:  .ColAlignment(3) = flexAlignLeftCenter:  .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 4) = "IdArea":    .ColWidth(4) = 0:
        .TextMatrix(0, 5) = "IdPer":     .ColWidth(5) = 0:
        GRID_COMBOLIST Fg1, 1     ' area
        GRID_COMBOLIST Fg1, 3     ' encargado
        .SelectionMode = flexSelectionByRow
    End With
   
    With Fg2 ' de la distribucion de las tareas
        .Rows = 2
        .Cols = 5
        .FixedRows = 2
        .RowHeight(0) = 250
        .RowHidden(1) = True
        .FrozenCols = 3
        .TextMatrix(0, 1) = "Tarea":        .ColWidth(1) = 4000:  .ColAlignment(1) = flexAlignLeftCenter:    .Row = 1: .Col = 1: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 2) = "U.M.":         .ColWidth(2) = 0:     .ColAlignment(2) = flexAlignCenterCenter:  .Row = 1: .Col = 2: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 3) = "Diverso":      .ColWidth(3) = 600:   .ColAlignment(3) = flexAlignCenterCenter:  .Row = 1: .Col = 3: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 4) = "IdTarea":      .ColWidth(4) = 0:
        .SelectionMode = flexSelectionByRow
        DoEvents
    End With
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarTareasFaltantes
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub pCargarTareasFaltantes()
    If QueHace = 3 Then Exit Sub
    Agregando = True
    Dim nSQL As String
    Dim nSQLNotIn As String
    Dim RstTmp As New ADODB.Recordset
    On Error GoTo error
    Me.MousePointer = vbHourglass
    
    ' generando la lista de id's ya seleccionados
    nSQLNotIn = GRID_GENERAR_SQL_ID(Fg2, 4, " AND pro_tareas.id ", "NOT IN", True)
    
    ' cargando las areas
    nSQL = "SELECT pro_tareas.descripcion, IIf([pro_tareas].[diverso]=0,'No','Si') AS esdiverso, mae_unidades.abrev, pro_tareas.id " _
        + vbCr + " FROM mae_unidades RIGHT JOIN pro_tareas ON mae_unidades.id = pro_tareas.idunimed " _
        + vbCr + " WHERE (((pro_tareas.id) Not In (SELECT pro_areadet.idtar FROM pro_areadet))) " & nSQLNotIn _
        + vbCr + " ORDER BY pro_tareas.diverso DESC , pro_tareas.descripcion;"
    
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    Agregando = True
    Do While Not RstTmp.EOF
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(RstTmp("descripcion"))
        Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosC(RstTmp("abrev"))
        Fg2.TextMatrix(Fg2.Rows - 1, 3) = NulosC(RstTmp("esdiverso"))
        Fg2.TextMatrix(Fg2.Rows - 1, 4) = NulosN(RstTmp("id"))
        RstTmp.MoveNext
    Loop
    
    Agregando = False
    If RstTmp.RecordCount <> 0 Then
        MsgBox "Se agregaron " & RstTmp.RecordCount & " Registros", vbInformation, xTitulo
    Else
        MsgBox "No hay registros para agregar", vbExclamation, xTitulo
    End If
    Set RstTmp = Nothing
    
    GRID_AGRUPAR Fg2, 4
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Set RstTmp = Nothing
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pCargarTareasFaltantes"
End Sub

'*****************************************************************************************************
'* Nombre           : pBuscarVSFlexGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pBuscarVSFlexGrid()
    On Error GoTo error
    If Me.TabOne1.CurrTab = 0 Then Exit Sub
    Dim xExport As New SGI2_funciones.formularios
    Dim xCampos(0, 3) As String
    'campo     'columna del grid    'tipo(N:Numerico, C:caracter, F:fecha)      campo predeterminado(0:no se muestra, -1:se muestra al iniciar el formulario)
    xCampos(0, 0) = "Tarea":        xCampos(0, 1) = "1":    xCampos(0, 2) = "C":    xCampos(0, 3) = "-1"
    xExport.VSFlexGrid_Buscar Me.hWnd, Fg2, xCampos()
    Set xExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pBuscarVSFlexGrid"
End Sub
