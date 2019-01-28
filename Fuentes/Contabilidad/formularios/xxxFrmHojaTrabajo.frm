VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmHojaTrabajo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Hoja de Trabajo"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6120
      Left            =   0
      TabIndex        =   11
      Top             =   945
      Width           =   11850
      _cx             =   20902
      _cy             =   10795
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
      Appearance      =   0
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "   Por Detalle   |    Por Cuenta    |  Por Sub Cuenta  "
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   0
      Position        =   1
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
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   5760
         Left            =   12765
         TabIndex        =   14
         Top             =   15
         Width           =   11820
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   5550
            Index           =   2
            Left            =   30
            TabIndex        =   17
            Top             =   105
            Width           =   11760
            _cx             =   20743
            _cy             =   9790
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
            Cols            =   21
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmHojaTrabajo.frx":0000
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   5760
         Left            =   12465
         TabIndex        =   13
         Top             =   15
         Width           =   11820
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   5550
            Index           =   1
            Left            =   30
            TabIndex        =   16
            Top             =   105
            Width           =   11760
            _cx             =   20743
            _cy             =   9790
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
            Cols            =   21
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmHojaTrabajo.frx":01A5
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   5760
         Left            =   15
         TabIndex        =   12
         Top             =   15
         Width           =   11820
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   5550
            Index           =   0
            Left            =   30
            TabIndex        =   15
            Top             =   105
            Width           =   11760
            _cx             =   20743
            _cy             =   9790
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
            Cols            =   21
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmHojaTrabajo.frx":034A
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
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   0
      TabIndex        =   2
      Top             =   7020
      Width           =   11865
      Begin VB.Label LblDescCuenta 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblDescCuenta"
         Height          =   300
         Left            =   1605
         TabIndex        =   4
         Top             =   165
         Width           =   10050
      End
      Begin VB.Label LbDescCuenta 
         Caption         =   "Cuenta Contable "
         Height          =   180
         Left            =   225
         TabIndex        =   3
         Top             =   210
         Width           =   1365
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Balance"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   14
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a Excel"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11055
      Top             =   135
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
            Picture         =   "FrmHojaTrabajo.frx":04EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo.frx":0A33
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo.frx":0DC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo.frx":0F1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo.frx":12B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo.frx":1435
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo.frx":1889
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo.frx":19A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo.frx":1EE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo.frx":2429
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo.frx":253D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo.frx":2651
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo.frx":2AA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo.frx":2C11
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   585
      Left            =   0
      TabIndex        =   6
      Top             =   300
      Width           =   11865
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   360
         Left            =   3510
         TabIndex        =   21
         Top             =   135
         Width           =   5235
         Begin VB.CommandButton cmd_periodo1 
            Height          =   240
            Left            =   2220
            Picture         =   "FrmHojaTrabajo.frx":3159
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   60
            Width           =   270
         End
         Begin VB.CommandButton cmd_periodo2 
            Height          =   240
            Left            =   4815
            Picture         =   "FrmHojaTrabajo.frx":34DB
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   60
            Width           =   270
         End
         Begin VB.Label Label5 
            Caption         =   "Per. Final"
            Height          =   195
            Left            =   2640
            TabIndex        =   27
            Top             =   75
            Width           =   720
         End
         Begin VB.Label Label4 
            Caption         =   "Per. Inicio"
            Height          =   195
            Left            =   0
            TabIndex        =   26
            Top             =   75
            Width           =   720
         End
         Begin VB.Label LblPerFin 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblPerFin"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3465
            TabIndex        =   25
            Top             =   30
            Width           =   1650
         End
         Begin VB.Label LblPerIni 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblPerIni"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   870
            TabIndex        =   24
            Top             =   30
            Width           =   1650
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   360
         Left            =   75
         TabIndex        =   18
         Top             =   150
         Width           =   2745
         Begin VB.OptionButton Option2 
            Caption         =   "Por Periodo"
            Height          =   195
            Left            =   1350
            TabIndex        =   20
            Top             =   90
            Width           =   1185
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Por Fecha"
            Height          =   195
            Left            =   105
            TabIndex        =   19
            Top             =   90
            Width           =   1185
         End
      End
      Begin VB.OptionButton OptDolares 
         Caption         =   "Dolares"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   9975
         TabIndex        =   8
         Top             =   240
         Width           =   900
      End
      Begin VB.OptionButton OptSoles 
         Caption         =   "Soles"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   9015
         TabIndex        =   7
         Top             =   240
         Width           =   900
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   4515
         TabIndex        =   0
         Top             =   195
         Width           =   1245
         _ExtentX        =   2196
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
         Valor           =   "25/04/2008"
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   300
         Left            =   6675
         TabIndex        =   1
         Top             =   195
         Width           =   1245
         _ExtentX        =   2196
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
         Valor           =   "25/04/2008"
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   2910
         X2              =   2910
         Y1              =   165
         Y2              =   525
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   2895
         X2              =   2895
         Y1              =   165
         Y2              =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Inicio"
         Height          =   195
         Left            =   3600
         TabIndex        =   10
         Top             =   225
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Final"
         Height          =   195
         Left            =   5880
         TabIndex        =   9
         Top             =   225
         Width           =   690
      End
   End
End
Attribute VB_Name = "FrmHojaTrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstTmp As New ADODB.Recordset
Dim SeEjecuto As Boolean

Sub Cargar(Indice As Integer)
    Dim Rst As New ADODB.Recordset
    Dim xFil As Integer
    Dim A As Integer
    
    PreparaRST_Tmp
    Fg1(Indice).Rows = 2
    
    'CARGANOS LOS MOVIMIENTOS DEL PERIODO ESPECIFICADO
    If Indice = 0 Then
        RST_Busq Rst, "SELECT con_planctas.id, con_planctas.iddes, con_planctas.iddes2, con_planctas.cuenta, con_planctas.descripcion, " _
            & " (SELECT Sum(IIf([impdebdol]=0,[impdebsol],[impdebdol]*[con_tc].[impven])) AS totdeb FROM con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
            & " WHERE (((con_diario.idcue)=con_planctas.id) AND ((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "')))) AS debe, " _
            & " (SELECT Sum(IIf([imphabdol]=0,[imphabsol],[imphabdol]*[con_tc].[impven])) AS totdeb FROM con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
            & " WHERE (((con_diario.idcue)=con_planctas.id) AND ((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "')))) AS haber " _
            & " From con_planctas " _
            & " WHERE ((((SELECT Sum(IIf([impdebdol]=0,[impdebsol],[impdebdol]*[con_tc].[impven])) AS totdeb FROM con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
            & " WHERE (((con_diario.idcue)=con_planctas.id) AND ((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "'))))) Is Not Null) " _
            & " AND (((SELECT Sum(IIf([imphabdol]=0,[imphabsol],[imphabdol]*[con_tc].[impven])) AS totdeb FROM con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
            & " WHERE (((con_diario.idcue)=con_planctas.id) AND ((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "'))))) Is Not Null)) " _
            & " ORDER BY con_planctas.cuenta", xCon
    End If
    If Indice = 1 Then
        'hoja de trabajo a 2 digitos
        RST_Busq Rst, "SELECT con_planctas_1.id, con_planctas_1.iddes, con_planctas_1.iddes2, con_planctas_1.cuenta, con_planctas_1.descripcion, " _
            & " (SELECT Sum(IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol]*[con_tc].[impven],[con_diario].[impdebsol])) AS impdeb " _
            & " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue " _
            & " WHERE (((con_planctas.cuenta) Like con_planctas_1.cuenta+'%') AND ((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') " _
            & " And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "')))) AS debe, " _
            & " (SELECT Sum(IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol]*[con_tc].[impven],[con_diario].[imphabsol])) AS impdeb " _
            & " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue " _
            & " WHERE (((con_planctas.cuenta) Like con_planctas_1.cuenta+ '%') AND ((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') " _
            & " And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "')))) AS haber " _
            & " FROM con_planctas AS con_planctas_1 WHERE ((((SELECT Sum(IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol]*[con_tc].[impven],[con_diario].[impdebsol])) AS impdeb" _
            & " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue " _
            & " WHERE (((con_planctas.cuenta) Like con_planctas_1.cuenta+'%') AND ((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') " _
            & " And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "'))))) Is Not Null) AND (((SELECT Sum(IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol]*[con_tc].[impven],[con_diario].[imphabsol])) AS impdeb " _
            & " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue " _
            & " WHERE (((con_planctas.cuenta) Like con_planctas_1.cuenta+ '%') AND ((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') " _
            & " And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "'))))) Is Not Null) AND ((Len([cuenta]))=2)) ORDER BY con_planctas_1.cuenta", xCon
    
    End If
    
    If Indice = 2 Then
        'hoja de trabajo a 3 digitos
        RST_Busq Rst, "SELECT con_planctas_1.id, con_planctas_1.iddes, con_planctas_1.iddes2, con_planctas_1.cuenta, con_planctas_1.descripcion, " _
            & " (SELECT Sum(IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol]*[con_tc].[impven],[con_diario].[impdebsol])) AS impdeb " _
            & " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue " _
            & " WHERE (((con_planctas.cuenta) Like con_planctas_1.cuenta+'%') AND ((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') " _
            & " And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "')))) AS debe, " _
            & " (SELECT Sum(IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol]*[con_tc].[impven],[con_diario].[imphabsol])) AS impdeb " _
            & " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue " _
            & " WHERE (((con_planctas.cuenta) Like con_planctas_1.cuenta+ '%') AND ((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') " _
            & " And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "')))) AS haber " _
            & " FROM con_planctas AS con_planctas_1 WHERE ((((SELECT Sum(IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol]*[con_tc].[impven],[con_diario].[impdebsol])) AS impdeb" _
            & " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue " _
            & " WHERE (((con_planctas.cuenta) Like con_planctas_1.cuenta+'%') AND ((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') " _
            & " And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "'))))) Is Not Null) AND (((SELECT Sum(IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol]*[con_tc].[impven],[con_diario].[imphabsol])) AS impdeb " _
            & " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue " _
            & " WHERE (((con_planctas.cuenta) Like con_planctas_1.cuenta+ '%') AND ((con_diario.fchasi)>=CDate('" & TxtFchIni.Valor & "') " _
            & " And (con_diario.fchasi)<=CDate('" & TxtFchFin.Valor & "'))))) Is Not Null) AND ((Len([cuenta]))=4)) ORDER BY con_planctas_1.cuenta", xCon

    End If
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            DoEvents
            RstTmp.AddNew
            RstTmp("id") = Rst("id")
            RstTmp("iddes") = Rst("iddes")
            RstTmp("iddes2") = NulosN(Rst("iddes2"))
            RstTmp("cuenta") = Rst("cuenta")
            RstTmp("descripcion") = Rst("descripcion")
            RstTmp("debe") = Rst("debe")
            RstTmp("haber") = Rst("haber")
            RstTmp.Update
            Rst.MoveNext
            
            If Rst.EOF = True Then Exit For
        Next A
    End If
    
    Set Rst = Nothing
    'cargamos los saldos del mes anterior
    If Indice = 0 Then
        RST_Busq Rst, "SELECT con_planctas.id, con_planctas.iddes, con_planctas.iddes2, con_planctas.cuenta, con_planctas.descripcion, " _
            & " (SELECT Sum(IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol]*[con_tc].[impven],[con_diario].[impdebsol])) AS impdeb FROM con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
            & " WHERE (((con_diario.idcue)=con_planctas.id) AND ((con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "') Or (con_diario.fchasi) Is Null))) AS debe, " _
            & " (SELECT Sum(IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol]*[con_tc].[impven],[con_diario].[imphabsol])) AS imphab FROM con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
            & " WHERE (((con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "') Or (con_diario.fchasi) Is Null) AND ((con_diario.idcue)=con_planctas.id))) AS haber " _
            & " From con_planctas " _
            & " WHERE ((((SELECT Sum(IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol]*[con_tc].[impven],[con_diario].[impdebsol])) AS impdeb " _
            & " FROM con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha WHERE (((con_diario.idcue)=con_planctas.id) " _
            & " AND ((con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "') Or (con_diario.fchasi) Is Null))))<>0 Or ((SELECT Sum(IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol]*[con_tc].[impven],[con_diario].[impdebsol])) AS impdeb " _
            & " FROM con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
            & " WHERE (((con_diario.idcue)=con_planctas.id) AND ((con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "') Or (con_diario.fchasi) Is Null)))) Is Not Null) " _
            & " AND (((SELECT Sum(IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol]*[con_tc].[impven],[con_diario].[imphabsol])) AS imphab " _
            & " FROM con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
            & " WHERE (((con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "') Or (con_diario.fchasi) Is Null) AND ((con_diario.idcue)=con_planctas.id))))<>0 " _
            & " Or ((SELECT Sum(IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol]*[con_tc].[impven],[con_diario].[imphabsol])) AS imphab " _
            & " FROM con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha WHERE (((con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "') Or (con_diario.fchasi) Is Null) " _
            & " AND ((con_diario.idcue)=con_planctas.id)))) Is Not Null)) ORDER BY con_planctas.cuenta", xCon
    End If
    If Indice = 1 Then
        RST_Busq Rst, "SELECT con_planctas_1.id, con_planctas_1.iddes, con_planctas_1.iddes2, con_planctas_1.cuenta, con_planctas_1.descripcion, " _
            & " (SELECT Sum(IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol]*[con_tc].[impven],[con_diario].[impdebsol])) AS impdeb " _
            & " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue " _
            & " WHERE (((con_planctas.cuenta) Like con_planctas_1.cuenta + '%') AND ((con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "') Or " _
            & " (con_diario.fchasi) Is Null))) AS debe, " _
            & " (SELECT Sum(IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol]*[con_tc].[impven],[con_diario].[imphabsol])) AS imphab " _
            & " FROM con_planctas  RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue " _
            & " WHERE (((con_planctas.cuenta)  Like con_planctas_1.cuenta + '%') AND ((con_tc.idmon)=2) AND ((con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "') " _
            & " or (con_diario.fchasi) Is Null))) AS haber " _
            & " FROM con_planctas AS con_planctas_1 WHERE ((((SELECT Sum(IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol]*[con_tc].[impven],[con_diario].[impdebsol])) AS impdeb " _
            & " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue " _
            & " WHERE (((con_planctas.cuenta) Like con_planctas_1.cuenta + '%') AND ((con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "') Or (con_diario.fchasi) Is Null)))) Is Not Null) " _
            & " AND (((SELECT Sum(IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol]*[con_tc].[impven],[con_diario].[imphabsol])) AS imphab FROM con_planctas  " _
            & " RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue WHERE (((con_planctas.cuenta) " _
            & " Like con_planctas_1.cuenta + '%') AND ((con_tc.idmon)=2) AND ((con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "') or (con_diario.fchasi) Is Null)))) Is Not Null) " _
            & " AND ((Len([cuenta]))=2)) ORDER BY con_planctas_1.cuenta", xCon
    End If
    
    If Indice = 2 Then
        RST_Busq Rst, "SELECT con_planctas_1.id, con_planctas_1.iddes, con_planctas_1.iddes2, con_planctas_1.cuenta, con_planctas_1.descripcion, " _
            & " (SELECT Sum(IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol]*[con_tc].[impven],[con_diario].[impdebsol])) AS impdeb " _
            & " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue " _
            & " WHERE (((con_planctas.cuenta) Like con_planctas_1.cuenta + '%') AND ((con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "') Or " _
            & " (con_diario.fchasi) Is Null))) AS debe, " _
            & " (SELECT Sum(IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol]*[con_tc].[impven],[con_diario].[imphabsol])) AS imphab " _
            & " FROM con_planctas  RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue " _
            & " WHERE (((con_planctas.cuenta)  Like con_planctas_1.cuenta + '%') AND ((con_tc.idmon)=2) AND ((con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "') " _
            & " or (con_diario.fchasi) Is Null))) AS haber " _
            & " FROM con_planctas AS con_planctas_1 WHERE ((((SELECT Sum(IIf([con_diario].[impdebdol]<>0,[con_diario].[impdebdol]*[con_tc].[impven],[con_diario].[impdebsol])) AS impdeb " _
            & " FROM con_planctas RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue " _
            & " WHERE (((con_planctas.cuenta) Like con_planctas_1.cuenta + '%') AND ((con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "') Or (con_diario.fchasi) Is Null)))) Is Not Null) " _
            & " AND (((SELECT Sum(IIf([con_diario].[imphabdol]<>0,[con_diario].[imphabdol]*[con_tc].[impven],[con_diario].[imphabsol])) AS imphab FROM con_planctas  " _
            & " RIGHT JOIN (con_diario LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON con_planctas.id = con_diario.idcue WHERE (((con_planctas.cuenta) " _
            & " Like con_planctas_1.cuenta + '%') AND ((con_tc.idmon)=2) AND ((con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "') or (con_diario.fchasi) Is Null)))) Is Not Null) " _
            & " AND ((Len([cuenta]))=4)) ORDER BY con_planctas_1.cuenta", xCon
    End If
      
   If Rst.RecordCount <> 0 Then
        If RstTmp.RecordCount <> 0 Then
            RstTmp.MoveFirst
        End If
        Rst.MoveFirst
        
        For A = 1 To Rst.RecordCount
            DoEvents
            If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
            RstTmp.Find "id = " & Rst("id") & ""
            If RstTmp.EOF = True Then
                RstTmp.AddNew
                RstTmp("id") = Rst("id")
                RstTmp("iddes") = Rst("iddes")
                RstTmp("iddes2") = NulosN(Rst("iddes2"))
                RstTmp("cuenta") = Rst("cuenta")
                RstTmp("descripcion") = Rst("descripcion")
                RstTmp("saldodeb") = Rst("debe")
                RstTmp("saldohab") = Rst("haber")
            Else
                RstTmp("saldodeb") = Rst("debe")
                RstTmp("saldohab") = Rst("haber")
                RstTmp.Update
            End If
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
    End If
    Set Rst = Nothing
    Set Rst = RstTmp
    Rst.Sort = "cuenta"
    
    xFil = 2
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            DoEvents
            Fg1(Indice).Rows = Fg1(Indice).Rows + 1
            Fg1(Indice).TextMatrix(xFil, 1) = Rst("cuenta")
            Fg1(Indice).TextMatrix(xFil, 2) = Rst("descripcion")
            Fg1(Indice).TextMatrix(xFil, 19) = Rst("iddes")
            Fg1(Indice).TextMatrix(xFil, 20) = NulosN(Rst("iddes2"))
            'Saldo anterior
            Fg1(Indice).TextMatrix(xFil, 3) = Format(Rst("saldodeb"), "0.00")
            Fg1(Indice).TextMatrix(xFil, 4) = Format(Rst("saldohab"), "0.00")
            
            'movimientos del ejercicio
            Fg1(Indice).TextMatrix(xFil, 5) = Format(Rst("debe"), "0.00")
            Fg1(Indice).TextMatrix(xFil, 6) = Format(Rst("haber"), "0.00")
            
            'sumas del mayor
            Fg1(Indice).TextMatrix(xFil, 7) = Format(Rst("debe") + Val(Fg1(Indice).TextMatrix(xFil, 3)), "0.00")
            Fg1(Indice).TextMatrix(xFil, 8) = Format(Rst("haber") + Val(Fg1(Indice).TextMatrix(xFil, 4)), "0.00")
            
            
            'saldo
            If Val(Fg1(Indice).TextMatrix(xFil, 7)) - Val(Fg1(Indice).TextMatrix(xFil, 8)) > 0 Then
                Fg1(Indice).TextMatrix(xFil, 9) = Val(Fg1(Indice).TextMatrix(xFil, 7)) - Val(Fg1(Indice).TextMatrix(xFil, 8))
                Fg1(Indice).TextMatrix(xFil, 9) = Format(Fg1(Indice).TextMatrix(xFil, 9), "0.00")
                Fg1(Indice).TextMatrix(xFil, 10) = "0.00"
            Else
                Fg1(Indice).TextMatrix(xFil, 9) = "0.00"
                Fg1(Indice).TextMatrix(xFil, 10) = Val(Fg1(Indice).TextMatrix(xFil, 8)) - Val(Fg1(Indice).TextMatrix(xFil, 7))
                Fg1(Indice).TextMatrix(xFil, 10) = Format(Fg1(Indice).TextMatrix(xFil, 10), "0.00")
            End If
            
            'DISTRIBUIMOS LAS CUENTAS
            'CUENTAS DEL BALANCE
            If Rst("iddes") = 1 Then
                Fg1(Indice).TextMatrix(xFil, 11) = Fg1(Indice).TextMatrix(xFil, 9)
                Fg1(Indice).TextMatrix(xFil, 12) = Fg1(Indice).TextMatrix(xFil, 10)
            End If
            
            'CUENTAS DE TRANSFERENCIA
            If Rst("iddes") = 4 Then
                Fg1(Indice).TextMatrix(xFil, 13) = Fg1(Indice).TextMatrix(xFil, 9)
                Fg1(Indice).TextMatrix(xFil, 14) = Fg1(Indice).TextMatrix(xFil, 10)
            End If
            
            'CUENTAS GANANCIA POR NATURALEZA
            If Rst("iddes") = 2 Then
                Fg1(Indice).TextMatrix(xFil, 15) = Fg1(Indice).TextMatrix(xFil, 9)
                Fg1(Indice).TextMatrix(xFil, 16) = Fg1(Indice).TextMatrix(xFil, 10)
            End If
            
            'CUENTAS GANANCIA POR FUNCION
            If Rst("iddes") = 3 Then
                Fg1(Indice).TextMatrix(xFil, 17) = Fg1(Indice).TextMatrix(xFil, 9)
                Fg1(Indice).TextMatrix(xFil, 18) = Fg1(Indice).TextMatrix(xFil, 10)
            End If
            
            'DISTRIBUIMOS LAS CUENTAS QUE DOBLETEAN EN LA HOJA DE TRABAJO (CUENTAS JUGADORAS)
            'CUENTAS DEL BALANCE
            If Rst("iddes2") = 1 Then
                Fg1(Indice).TextMatrix(xFil, 11) = Fg1(Indice).TextMatrix(xFil, 9)
                Fg1(Indice).TextMatrix(xFil, 12) = Fg1(Indice).TextMatrix(xFil, 10)
            End If
            
            'CUENTAS DE TRANSFERENCIA
            If Rst("iddes2") = 4 Then
                Fg1(Indice).TextMatrix(xFil, 13) = Fg1(Indice).TextMatrix(xFil, 9)
                Fg1(Indice).TextMatrix(xFil, 14) = Fg1(Indice).TextMatrix(xFil, 10)
            End If
            
            'CUENTAS GANANCIA POR NATURALEZA
            If Rst("iddes2") = 2 Then
                Fg1(Indice).TextMatrix(xFil, 15) = Fg1(Indice).TextMatrix(xFil, 9)
                Fg1(Indice).TextMatrix(xFil, 16) = Fg1(Indice).TextMatrix(xFil, 10)
            End If
            
            'CUENTAS GANANCIA POR FUNCION
            If Rst("iddes2") = 3 Then
                Fg1(Indice).TextMatrix(xFil, 17) = Fg1(Indice).TextMatrix(xFil, 9)
                Fg1(Indice).TextMatrix(xFil, 18) = Fg1(Indice).TextMatrix(xFil, 10)
            End If
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
            xFil = xFil + 1
        Next A
    Else
        MsgBox "No hay registros para procesar la hoja de trabajo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
    
    
End Sub

Sub Totalizar(Indice As Integer)
    Dim A As Integer
    Dim xTotal1, xTotal2, xTotal3, xTotal4, xTotal5, xTotal6, xTotal7, xTotal8, xTotal9, xTotal10  As Double
    Dim xTotal11, xTotal12, xTotal13, xTotal14, xTotal15, xTotal16 As Double
    
    For A = 2 To Fg1(Indice).Rows - 1
        DoEvents
        xTotal1 = xTotal1 + Val(Fg1(Indice).TextMatrix(A, 3))
        xTotal2 = xTotal2 + Val(Fg1(Indice).TextMatrix(A, 4))
        xTotal3 = xTotal3 + Val(Fg1(Indice).TextMatrix(A, 5))
        xTotal4 = xTotal4 + Val(Fg1(Indice).TextMatrix(A, 6))
        xTotal5 = xTotal5 + Val(Fg1(Indice).TextMatrix(A, 7))
        xTotal6 = xTotal6 + Val(Fg1(Indice).TextMatrix(A, 8))
        xTotal7 = xTotal7 + Val(Fg1(Indice).TextMatrix(A, 9))
        xTotal8 = xTotal8 + Val(Fg1(Indice).TextMatrix(A, 10))
        xTotal9 = xTotal9 + Val(Fg1(Indice).TextMatrix(A, 11))
        xTotal10 = xTotal10 + Val(Fg1(Indice).TextMatrix(A, 12))
        xTotal11 = xTotal11 + Val(Fg1(Indice).TextMatrix(A, 13))
        xTotal12 = xTotal12 + Val(Fg1(Indice).TextMatrix(A, 14))
        xTotal13 = xTotal13 + Val(Fg1(Indice).TextMatrix(A, 15))
        xTotal14 = xTotal14 + Val(Fg1(Indice).TextMatrix(A, 16))
        xTotal15 = xTotal15 + Val(Fg1(Indice).TextMatrix(A, 17))
        xTotal16 = xTotal16 + Val(Fg1(Indice).TextMatrix(A, 18))
    Next A
    
    Fg1(Indice).Rows = Fg1(Indice).Rows + 1
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 2) = "T O T A L E S ==>"
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 3) = Format(xTotal1, "0.00")
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 4) = Format(xTotal2, "0.00")
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 5) = Format(xTotal3, "0.00")
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 6) = Format(xTotal4, "0.00")
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 7) = Format(xTotal5, "0.00")
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 8) = Format(xTotal6, "0.00")
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 9) = Format(xTotal7, "0.00")
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 10) = Format(xTotal8, "0.00")
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 11) = Format(xTotal9, "0.00")
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 12) = Format(xTotal10, "0.00")
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 13) = Format(xTotal11, "0.00")
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 14) = Format(xTotal12, "0.00")
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 15) = Format(xTotal13, "0.00")
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 16) = Format(xTotal14, "0.00")
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 17) = Format(xTotal15, "0.00")
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 18) = Format(xTotal16, "0.00")

    Fg1(Indice).Rows = Fg1(Indice).Rows + 1
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 2) = "R E S U L T A D O ==>"
        
    If Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 10)) - Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 9)) > 0 Then
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 9) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 9)) - Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 10))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 9) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 9), "0.00")
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 10) = "0.00"
    Else
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 9) = "0.00"
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 10) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 10)) - Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 9))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 10) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 10), "0.00")
    End If

    If Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 12)) - Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 11)) > 0 Then
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 11) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 12)) - Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 11))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 11) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 11), "0.00")
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 12) = "0.00"
    Else
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 11) = "0.00"
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 12) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 11)) - Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 12))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 12) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 12), "0.00")
    End If
    
    If Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 14)) - Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 13)) > 0 Then
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 13) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 14)) - Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 13))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 13) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 13), "0.00")
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 14) = "0.00"
    Else
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 13) = "0.00"
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 14) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 13)) - Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 14))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 14) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 14), "0.00")
    End If
    
    If Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 16)) - Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 15)) > 0 Then
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 15) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 16)) - Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 15))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 15) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 15), "0.00")
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 16) = "0.00"
    Else
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 15) = "0.00"
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 16) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 15)) - Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 16))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 16) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 16), "0.00")
    End If
    
    If Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 18)) - Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 17)) > 0 Then
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 17) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 18)) - Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 17))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 17) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 17), "0.00")
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 18) = "0.00"
    Else
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 17) = "0.00"
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 18) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 17)) - Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 18))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 18) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 18), "0.00")
    End If
    
    Fg1(Indice).Rows = Fg1(Indice).Rows + 1
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 2) = "S U M A S  I G U A L E S ==>"
    
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 9) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 9)) + Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 9))
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 10) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 10)) + Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 10))
    
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 11) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 11)) + Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 11))
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 12) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 12)) + Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 12))

    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 13) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 13)) + Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 13))
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 14) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 14)) + Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 14))

    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 15) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 15)) + Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 15))
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 16) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 16)) + Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 16))

    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 17) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 17)) + Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 17))
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 18) = Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 18)) + Val(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 18))
End Sub

Private Sub CmdImprimir_Click()

End Sub

Private Sub CmdMuestra_Click()
    If TxtFchIni.Valor = "" Then
        MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    If TxtFchFin.Valor = "" Then
        MsgBox "No ha especificado la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchFin.SetFocus
        Exit Sub
    End If
    
    If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
    MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    Procesar
End Sub

Sub Procesar()
    Dim A As Integer
    
    For A = 0 To 0
        DoEvents
        Cargar A
        Totalizar A
    
        Fg1(A).FrozenCols = 2
        
        With Fg1(A)
            'AMARILLO
            .Select 2, 1, Fg1(A).Rows - 1, 2
            .FillStyle = flexFillRepeat
            .CellBackColor = &HECFFFF
        
            'AMARILLO
            .Select 2, 5, Fg1(A).Rows - 1, 6
            .FillStyle = flexFillRepeat
            .CellBackColor = &HECFFFF
        
            'AMARILLO
            .Select 2, 9, Fg1(A).Rows - 1, 10
            .FillStyle = flexFillRepeat
            .CellBackColor = &HECFFFF
        
            'AMARILLO
            .Select 2, 13, Fg1(A).Rows - 1, 14
            .FillStyle = flexFillRepeat
            .CellBackColor = &HECFFFF
            
            'AMARILLO
            .Select 2, 17, Fg1(A).Rows - 1, 18
            .FillStyle = flexFillRepeat
            .CellBackColor = &HECFFFF
            
            .Select Fg1(A).Rows - 3, 1, Fg1(A).Rows - 1, Fg1(A).Cols - 1
            .FillStyle = flexFillRepeat
            .CellBackColor = &HE0FEE7
            
            .Select 2, 1, 2, 1
        End With
    Next A
    TabOne1.CurrTab = 0
    LblDescCuenta.Caption = Fg1(0).TextMatrix(2, 2)
End Sub

Private Sub cmd_periodo1_Click()
    Dim xMesIni As Integer
    Dim xFchIni As String
    xMesIni = SeleccionaMes(xCon)
    LblPerIni.Caption = Busca_Codigo(xMesIni, "id", "descripcion", "con_meses", "N", xCon)
    TxtFchIni.Valor = "01/" + Format(xMesIni, "00") + "/" + Format(AnoTra, "0000")
End Sub

Private Sub cmd_periodo2_Click()
    Dim xMesIni, NumDias As Integer
    Dim xxFchIni As String
    
    xMesIni = SeleccionaMes(xCon)
    LblPerFin.Caption = Busca_Codigo(xMesIni, "id", "descripcion", "con_meses", "N", xCon)
    
    xxFchIni = "01/" + Format(xMesIni, "00") + "/" + Format(AnoTra, "0000")
    NumDias = HallaDiasMes(CDate(xxFchIni))
    TxtFchFin.Valor = Format(NumDias, "00") + "/" + Format(xMesIni, "00") + "/" + Format(AnoTra, "0000")
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        Setea
        OptSoles.Value = True
        Option1.Value = True
        Option1_Click
    End If
End Sub

Sub Setea()
    'usamos la columna 19 para almacenar el destino de cada cuenta en la hoja de trabajo
    Dim A As Integer
    
    For A = 0 To 2
         Fg1(A).ColWidth(19) = 0
         Fg1(A).ColWidth(20) = 0
         Fg1(A).Rows = 2
         Fg1(A).TextMatrix(0, 1) = "          1"
         Fg1(A).TextMatrix(1, 1) = "          1"
         Fg1(A).TextMatrix(0, 1) = "Nº Cuenta"
         Fg1(A).TextMatrix(1, 1) = "Nº Cuenta"
         Fg1(A).TextMatrix(0, 2) = "Descripcion"
         Fg1(A).TextMatrix(1, 2) = "Descripcion"
         
         'Fg1.MergeCells = flexMergeFree
         Fg1(A).Redraw = False
         Fg1(A).MergeCol(0) = True
         Fg1(A).MergeCol(1) = True
         Fg1(A).MergeCol(2) = True
         
         Fg1(A).MergeCells = 2
         Fg1(A).Redraw = True
         
         With Fg1(A)
             .MergeCells = flexMergeFree
             .MergeRow(-1) = True
             .Cell(flexcpText, 0, 3, 0, 4) = "Saldos Iniciales"
             .Cell(flexcpText, 0, 5, 0, 6) = "Movimiento del Periodo"
             .Cell(flexcpText, 0, 7, 0, 8) = "Sumas del Mayor"
             .Cell(flexcpText, 0, 9, 0, 10) = "Saldos Al"
             .Cell(flexcpText, 0, 11, 0, 12) = "Cuentas del Pasivo"
             .Cell(flexcpText, 0, 13, 0, 14) = "Transferencias y Canc."
             .Cell(flexcpText, 0, 15, 0, 16) = "Resultados x Naturaleza"
             .Cell(flexcpText, 0, 17, 0, 18) = "Resultados x Funcion"
             .Cell(flexcpBackColor, 0, 0, Fg1(A).Rows - 1, Fg1(A).Cols - 1) = &H8000000F
         End With
        
         Fg1(A).ColWidth(3) = 1100
         Fg1(A).ColWidth(4) = 1100
         Fg1(A).ColWidth(5) = 1100
         Fg1(A).ColWidth(6) = 1100
         Fg1(A).ColWidth(7) = 1100
         Fg1(A).ColWidth(8) = 1100
         Fg1(A).ColWidth(9) = 1100
         Fg1(A).ColWidth(10) = 1100
         Fg1(A).ColWidth(11) = 1100
         Fg1(A).ColWidth(12) = 1100
         Fg1(A).ColWidth(13) = 1100
         Fg1(A).ColWidth(14) = 1100
         Fg1(A).ColWidth(15) = 1100
         Fg1(A).ColWidth(16) = 1100
         Fg1(A).ColWidth(17) = 1100
         Fg1(A).ColWidth(18) = 1100
             
         Fg1(A).TextMatrix(1, 3) = "Debe"
         Fg1(A).TextMatrix(1, 4) = "Haber"
         Fg1(A).TextMatrix(1, 5) = "Debe"
         Fg1(A).TextMatrix(1, 6) = "Haber"
         Fg1(A).TextMatrix(1, 7) = "Debe"
         Fg1(A).TextMatrix(1, 8) = "Haber"
         Fg1(A).TextMatrix(1, 9) = "Debe"
         Fg1(A).TextMatrix(1, 10) = "Haber"
         Fg1(A).TextMatrix(1, 11) = "Debe"
         Fg1(A).TextMatrix(1, 12) = "Haber"
         Fg1(A).TextMatrix(1, 13) = "Debe"
         Fg1(A).TextMatrix(1, 14) = "Haber"
         Fg1(A).TextMatrix(1, 15) = "Debe"
         Fg1(A).TextMatrix(1, 16) = "Haber"
         Fg1(A).TextMatrix(1, 17) = "Debe"
         Fg1(A).TextMatrix(1, 18) = "Haber"
    Next A
End Sub

Sub ExportarComprasExcel(Indice As Integer)
    If Fg1(Indice).Rows = 1 Then
        MsgBox "No se ha registrado compras para exportar", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
        Exit Sub
    End If
    
    Dim A As Integer
    Dim B As Integer
    Dim xFilas As Integer
    
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    
    objExcel.Visible = True
    objExcel.SheetsInNewWorkbook = 1
    
    'abre el Libro
    objExcel.WindowState = 1
    objExcel.Workbooks.Add
    
    
    With objExcel.ActiveSheet
        .Cells(1, 2) = NomEmp
        .Cells(1, 13) = Date
        .Cells(2, 2) = "Nº R.U.C. : " + NumRUC
        
        xFilas = 4
        For A = 0 To 0
            For B = 1 To Fg1(Indice).Cols - 1
                If B = 1 Or B = 2 Then
                    .Cells(xFilas, B + 1) = "'" + Fg1(Indice).TextMatrix(A, B)
                Else
                    If B = 3 Or B = 5 Or B = 7 Or B = 9 Or B = 11 Or B = 13 Or B = 15 Or B = 17 Then
                        .Cells(xFilas, B + 1) = "'" + Fg1(Indice).TextMatrix(A, B)
                    End If
                End If
            Next B
            xFilas = xFilas + 1
        Next A
        
        For A = 1 To 1
            For B = 1 To Fg1(Indice).Cols - 1
                .Cells(xFilas, B + 1) = "'" + Fg1(Indice).TextMatrix(A, B)
            Next B
            xFilas = xFilas + 1
        Next A
        
        For A = 2 To Fg1(Indice).Rows - 1
            For B = 1 To Fg1(Indice).Cols - 1
                If B <= 2 Then
                    .Cells(xFilas, B + 1) = "'" + Fg1(Indice).TextMatrix(A, B)
                Else
                    .Cells(xFilas, B + 1) = Val(Fg1(Indice).TextMatrix(A, B))
                End If
            Next B
            xFilas = xFilas + 1
        Next A
    End With
    
    MsgBox "El proceso de exportacion termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Registro de Compras y Ventas"
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
End Sub


Private Sub Form_Load()
    TxtFchIni.Valor = Date
    TxtFchFin.Valor = Date
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    
    Frame3.BackColor = &H8000000F
    Frame4.BackColor = &H8000000F
    Frame5.BackColor = &H8000000F
    Frame6.BackColor = &H8000000F
    Frame7.BackColor = &H8000000F
    SeEjecuto = False
    LblDescCuenta.Caption = ""
    
    LblPerIni.Caption = ""
    LblPerFin.Caption = ""
    
    TabOne1.CurrTab = 0
End Sub

Private Sub Option1_Click()
    Frame7.Visible = False
End Sub

Private Sub Option2_Click()
    LblPerIni.Caption = ""
    LblPerFin.Caption = ""
    
    Frame7.Left = 3600
    Frame7.Top = 165
    Frame7.Visible = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        If TxtFchIni.Valor = "" Then
            MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchIni.SetFocus
            Exit Sub
        End If
        
        If TxtFchFin.Valor = "" Then
            MsgBox "No ha especificado la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchFin.SetFocus
            Exit Sub
        End If
        
        If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
        MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchIni.SetFocus
            Exit Sub
        End If
        Procesar
    End If
    
    If Button.Index = 3 Then
        ExportarComprasExcel TabOne1.CurrTab
    End If
    
    If Button.Index = 5 Then
        Unload Me
    End If
End Sub

Sub PreparaRST_Tmp()
    'Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(9, 3) As String

    xCampos(0, 0) = "id":           xCampos(0, 1) = "N":      xCampos(0, 2) = "20"
    xCampos(1, 0) = "iddes":        xCampos(1, 1) = "N":      xCampos(1, 2) = "200"
    xCampos(2, 0) = "iddes2":       xCampos(2, 1) = "N":      xCampos(2, 2) = "200"
    xCampos(3, 0) = "cuenta":       xCampos(3, 1) = "C":      xCampos(3, 2) = "15"
    xCampos(4, 0) = "descripcion":  xCampos(4, 1) = "C":      xCampos(4, 2) = "100"
    xCampos(5, 0) = "debe":         xCampos(5, 1) = "D":      xCampos(5, 2) = "200"
    xCampos(6, 0) = "haber":        xCampos(6, 1) = "D":      xCampos(6, 2) = "200"
    xCampos(7, 0) = "saldodeb":     xCampos(7, 1) = "D":      xCampos(7, 2) = "200"
    xCampos(8, 0) = "saldohab":     xCampos(8, 1) = "D":      xCampos(8, 2) = "200"
    
    Set RstTmp = xFun.CrearRstTMP(xCampos)

    RstTmp.Open
End Sub


Sub ImprimirBalance()
'    Set RstTmp = Nothing
'    PreparaRST_Tmp
'    Dim A As Integer
'
'    For A = 2 To Fg1.Rows - 1
'        If NulosN(Fg1.TextMatrix(A, 19)) = 1 Then
'            RstTmp.AddNew
'            RstTmp("numcue") = Fg1.TextMatrix(A, 1)
'            RstTmp("descripcion") = Fg1.TextMatrix(A, 2)
'            RstTmp("debe") = NulosN(Fg1.TextMatrix(A, 11))
'            RstTmp("haber") = NulosN(Fg1.TextMatrix(A, 12))
'            RstTmp.Update
'        End If
'    Next A
'
'    For A = 2 To Fg1.Rows - 1
'        If NulosN(Fg1.TextMatrix(A, 20)) = 1 Then
'            RstTmp.AddNew
'            RstTmp("numcue") = Fg1.TextMatrix(A, 1)
'            RstTmp("descripcion") = Fg1.TextMatrix(A, 2)
'            RstTmp("debe") = NulosN(Fg1.TextMatrix(A, 11))
'            RstTmp("haber") = NulosN(Fg1.TextMatrix(A, 12))
'            RstTmp.Update
'        End If
'    Next A
'
'    RstTmp.Sort = "numcue"
'
'    RptBalance.Sections("Sección2").Controls("txtempresa").Caption = NomEmp
'    RptBalance.Sections("Sección2").Controls("txtnumruc").Caption = NumRUC
'    RptBalance.Sections("Sección2").Controls("txtfchemi").Caption = Date
'    RptBalance.Sections("Sección2").Controls("txttitulo").Caption = "BALANCE GENERAL"
'
'    RptBalance.Sections("Sección3").Controls("txttotdebe").Caption = Format(Fg1.TextMatrix(Fg1.Rows - 3, 11), "0.00")
'    RptBalance.Sections("Sección3").Controls("txttothaber").Caption = Format(Fg1.TextMatrix(Fg1.Rows - 3, 12), "0.00")
'
'    RptBalance.Sections("Sección3").Controls("txttotdebe2").Caption = Format(Fg1.TextMatrix(Fg1.Rows - 2, 11), "0.00")
'    RptBalance.Sections("Sección3").Controls("txttothaber2").Caption = Format(Fg1.TextMatrix(Fg1.Rows - 2, 12), "0.00")
'
'    RptBalance.Sections("Sección3").Controls("txttotdebe3").Caption = Format(Fg1.TextMatrix(Fg1.Rows - 1, 11), "0.00")
'    RptBalance.Sections("Sección3").Controls("txttothaber3").Caption = Format(Fg1.TextMatrix(Fg1.Rows - 1, 12), "0.00")
'
'    Set RptBalance.DataSource = RstTmp
'    RptBalance.Width = 11955
'    RptBalance.Height = 7965
'
'    'RptBalance.Orientation = rptOrientLandscape
'    RptBalance.Show vbModal
End Sub

'Sub ImprimirGananciaNaturaleza()
'    Set RstTmp = Nothing
'    PreparaRST_Tmp
'    Dim A As Integer
'
'    For A = 2 To Fg1.Rows - 1
'        If NulosN(Fg1.TextMatrix(A, 19)) = 2 Then
'            RstTmp.AddNew
'            RstTmp("numcue") = Fg1.TextMatrix(A, 1)
'            RstTmp("descripcion") = Fg1.TextMatrix(A, 2)
'            RstTmp("debe") = NulosN(Fg1.TextMatrix(A, 15))
'            RstTmp("haber") = NulosN(Fg1.TextMatrix(A, 16))
'            RstTmp.Update
'        End If
'    Next A
'
'    For A = 2 To Fg1.Rows - 1
'        If NulosN(Fg1.TextMatrix(A, 20)) = 2 Then
'            RstTmp.AddNew
'            RstTmp("numcue") = Fg1.TextMatrix(A, 1)
'            RstTmp("descripcion") = Fg1.TextMatrix(A, 2)
'            RstTmp("debe") = NulosN(Fg1.TextMatrix(A, 15))
'            RstTmp("haber") = NulosN(Fg1.TextMatrix(A, 16))
'            RstTmp.Update
'        End If
'    Next A
'
'    RstTmp.Sort = "numcue"
'
'    RptBalance.Sections("Sección2").Controls("txtempresa").Caption = NomEmp
'    RptBalance.Sections("Sección2").Controls("txtnumruc").Caption = NumRUC
'    RptBalance.Sections("Sección2").Controls("txtfchemi").Caption = Date
'    RptBalance.Sections("Sección2").Controls("txttitulo").Caption = "RESULTADO POR NATURALEZA"
'
'    RptBalance.Sections("Sección3").Controls("txttotdebe").Caption = Format(Fg1.TextMatrix(Fg1.Rows - 3, 15), "0.00")
'    RptBalance.Sections("Sección3").Controls("txttothaber").Caption = Format(Fg1.TextMatrix(Fg1.Rows - 3, 16), "0.00")
'
'    RptBalance.Sections("Sección3").Controls("txttotdebe2").Caption = Format(Fg1.TextMatrix(Fg1.Rows - 2, 15), "0.00")
'    RptBalance.Sections("Sección3").Controls("txttothaber2").Caption = Format(Fg1.TextMatrix(Fg1.Rows - 2, 16), "0.00")
'
'    RptBalance.Sections("Sección3").Controls("txttotdebe3").Caption = Format(Fg1.TextMatrix(Fg1.Rows - 1, 15), "0.00")
'    RptBalance.Sections("Sección3").Controls("txttothaber3").Caption = Format(Fg1.TextMatrix(Fg1.Rows - 1, 16), "0.00")
'
'    Set RptBalance.DataSource = RstTmp
'    RptBalance.Width = 11955
'    RptBalance.Height = 7965
'    RptBalance.Show vbModal
'End Sub

'Sub ImprimirGananciaFuncion()
'    Set RstTmp = Nothing
'    PreparaRST_Tmp
'    Dim A As Integer
'
'    For A = 2 To Fg1.Rows - 1
'        If NulosN(Fg1.TextMatrix(A, 19)) = 3 Then
'            RstTmp.AddNew
'            RstTmp("numcue") = Fg1.TextMatrix(A, 1)
'            RstTmp("descripcion") = Fg1.TextMatrix(A, 2)
'            RstTmp("debe") = NulosN(Fg1.TextMatrix(A, 17))
'            RstTmp("haber") = NulosN(Fg1.TextMatrix(A, 18))
'            RstTmp.Update
'        End If
'    Next A
'
'    For A = 2 To Fg1.Rows - 1
'        If NulosN(Fg1.TextMatrix(A, 20)) = 3 Then
'            RstTmp.AddNew
'            RstTmp("numcue") = Fg1.TextMatrix(A, 1)
'            RstTmp("descripcion") = Fg1.TextMatrix(A, 2)
'            RstTmp("debe") = NulosN(Fg1.TextMatrix(A, 17))
'            RstTmp("haber") = NulosN(Fg1.TextMatrix(A, 18))
'            RstTmp.Update
'        End If
'    Next A
'
'    RstTmp.Sort = "numcue"
'
'    RptBalance.Sections("Sección2").Controls("txtempresa").Caption = NomEmp
'    RptBalance.Sections("Sección2").Controls("txtnumruc").Caption = NumRUC
'    RptBalance.Sections("Sección2").Controls("txtfchemi").Caption = Date
'    RptBalance.Sections("Sección2").Controls("txttitulo").Caption = "RESULTADO POR FUNCION"
'
'    RptBalance.Sections("Sección3").Controls("txttotdebe").Caption = Format(Fg1.TextMatrix(Fg1.Rows - 3, 17), "0.00")
'    RptBalance.Sections("Sección3").Controls("txttothaber").Caption = Format(Fg1.TextMatrix(Fg1.Rows - 3, 18), "0.00")
'
'    RptBalance.Sections("Sección3").Controls("txttotdebe2").Caption = Format(Fg1.TextMatrix(Fg1.Rows - 2, 17), "0.00")
'    RptBalance.Sections("Sección3").Controls("txttothaber2").Caption = Format(Fg1.TextMatrix(Fg1.Rows - 2, 18), "0.00")
'
'    RptBalance.Sections("Sección3").Controls("txttotdebe3").Caption = Format(Fg1.TextMatrix(Fg1.Rows - 1, 17), "0.00")
'    RptBalance.Sections("Sección3").Controls("txttothaber3").Caption = Format(Fg1.TextMatrix(Fg1.Rows - 1, 18), "0.00")
'
'    Set RptBalance.DataSource = RstTmp
'    RptBalance.Width = 11955
'    RptBalance.Height = 7965
'    RptBalance.Show vbModal
'End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then
        ImprimirBalance
    End If
    If ButtonMenu.Index = 2 Then
        'ImprimirGananciaNaturaleza
    End If
    If ButtonMenu.Index = 3 Then
        'ImprimirGananciaFuncion
    End If
End Sub
