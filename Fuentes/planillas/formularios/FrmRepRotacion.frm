VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmRepRotacion 
   Caption         =   "Producción - Reporte de Asistencia"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   9600
      TabIndex        =   21
      Top             =   7800
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   225
         TabIndex        =   22
         Top             =   465
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblproc 
         AutoSize        =   -1  'True
         Caption         =   "lblproc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   225
         TabIndex        =   23
         Top             =   180
         Width           =   570
      End
      Begin VB.Shape Shape1 
         Height          =   765
         Index           =   0
         Left            =   60
         Top             =   90
         Width           =   5805
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   5320
      Index           =   0
      Left            =   30
      TabIndex        =   10
      Top             =   7800
      Visible         =   0   'False
      Width           =   9530
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   9260
         Picture         =   "FrmRepRotacion.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   11
         ToolTipText     =   "Cerrar"
         Top             =   30
         Width           =   195
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   4270
         Index           =   2
         Left            =   90
         TabIndex        =   13
         Top             =   870
         Width           =   9325
         _cx             =   16448
         _cy             =   7532
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
         Rows            =   2
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmRepRotacion.frx":02EC
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
         Editable        =   2
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
      Begin VB.Label lblDetalle 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDetalle"
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   2820
         TabIndex        =   17
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Detalle"
         Height          =   195
         Index           =   19
         Left            =   2220
         TabIndex        =   16
         Top             =   390
         Width           =   495
      End
      Begin VB.Label lblDia 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblDia"
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   510
         TabIndex        =   15
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Dia"
         Height          =   195
         Index           =   18
         Left            =   150
         TabIndex        =   14
         Top             =   420
         Width           =   240
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de Reporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   28
         Left            =   105
         TabIndex        =   12
         Top             =   60
         Width           =   1620
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   5
         X1              =   30
         X2              =   9500
         Y1              =   5290
         Y2              =   5290
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   4
         X1              =   0
         X2              =   9595
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   5
         X1              =   9500
         X2              =   9500
         Y1              =   0
         Y2              =   5290
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   30
         Top             =   30
         Width           =   9440
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11040
      Top             =   90
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
            Picture         =   "FrmRepRotacion.frx":03CA
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepRotacion.frx":090E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepRotacion.frx":0CA0
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepRotacion.frx":0E24
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepRotacion.frx":1278
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepRotacion.frx":1390
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepRotacion.frx":18D4
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepRotacion.frx":1E18
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepRotacion.frx":1F2C
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepRotacion.frx":2040
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepRotacion.frx":2494
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepRotacion.frx":2600
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepRotacion.frx":2B48
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
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Solicitud de Materiales"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Solicitud de Linea"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame2 
      Height          =   7500
      Left            =   0
      TabIndex        =   1
      Top             =   250
      Width           =   11820
      Begin VB.Frame Frame1 
         Caption         =   "[ Agrupamiento ]"
         Height          =   615
         Left            =   3930
         TabIndex        =   18
         Top             =   90
         Width           =   3165
         Begin VB.OptionButton OptAgrupar 
            Caption         =   "Agrupar x Cargo"
            Height          =   195
            Index           =   1
            Left            =   1620
            TabIndex        =   20
            Top             =   270
            Width           =   1515
         End
         Begin VB.OptionButton OptAgrupar 
            Caption         =   "Agrupar x Area"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   19
            Top             =   270
            Width           =   1515
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "[ Fecha ]"
         Height          =   615
         Left            =   60
         TabIndex        =   5
         Top             =   90
         Width           =   3825
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchAsisIni 
            Height          =   300
            Left            =   630
            TabIndex        =   6
            Top             =   225
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
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchAsisFin 
            Height          =   300
            Left            =   2490
            TabIndex        =   7
            Top             =   225
            Width           =   1230
            _ExtentX        =   2170
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
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Inicio"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   285
            Width           =   375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fin"
            Height          =   195
            Left            =   2130
            TabIndex        =   8
            Top             =   285
            Width           =   210
         End
      End
      Begin SizerOneLibCtl.TabOne TabOne1 
         Height          =   6690
         Left            =   30
         TabIndex        =   2
         Top             =   735
         Width           =   11745
         _cx             =   20717
         _cy             =   11800
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
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
         Caption         =   "     Detalle    |    Resumen    "
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   0
         Position        =   2
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   -1  'True
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         Begin VSFlex7Ctl.VSFlexGrid fg 
            Height          =   6660
            Index           =   0
            Left            =   330
            TabIndex        =   3
            Top             =   15
            Width           =   11400
            _cx             =   20108
            _cy             =   11747
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
            Rows            =   2
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmRepRotacion.frx":2EDA
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
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
         Begin VSFlex7Ctl.VSFlexGrid fg 
            Height          =   6660
            Index           =   1
            Left            =   12675
            TabIndex        =   4
            Top             =   15
            Width           =   11400
            _cx             =   20108
            _cy             =   11747
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
            Rows            =   2
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmRepRotacion.frx":2F53
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
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
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu00 
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "FrmRepRotacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rst As New ADODB.Recordset
Dim RstOrd As New ADODB.Recordset
Dim Quehace As Integer
Dim SeEjecuto As Boolean
Dim Agregando As Boolean               ' para saber cuando se este agregando FILAS AL CONTROL grid de productos
Dim IdMenuActivo As Integer            ' INDICA EL CODIGO DEL MENU ACTIVO
Dim agregados As Integer
Dim xHorIni As Date                    ' ESPECIFICA LA HORA DE INICIO
Dim mCorrelativo As Long               ' para diferenciar la fecha de entrega del pedido cuando se necesite modificar
Dim mIdRegistro&                       ' identificador del registro
Dim mMesActivo As Integer              ' indica el mes activo

Dim cSQL As String
Dim CONS_FECH_ASISTENCIA As String
Dim CONS_HORA_ASISTENCIA As String
Dim cPERSONAL As String
Dim CAREA As String
Dim CCARGO As String
Dim xTitulo As String
Dim CALCULANDO_ As Boolean

Dim xRsIng As New ADODB.Recordset
Dim xRsSal As New ADODB.Recordset
Dim xRsRei As New ADODB.Recordset
Dim xRsPers As New ADODB.Recordset
Dim xRsPers2 As New ADODB.Recordset
Dim xRsCriAgr As New ADODB.Recordset

Dim OrigFX As Long
Dim OrigFY As Long

Private Function hallarPersonal(FECH_ As String) As ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    
    cSQL = "SELECT pla_periodolaboral.idemp, pla_empleados.idarea, pla_empleados.idcargo, pla_empleados.nombre " _
        + vbCr + "FROM (pla_periodolaboral LEFT JOIN pla_empleados ON pla_periodolaboral.idemp = pla_empleados.id) LEFT JOIN pla_recmarcacion ON pla_periodolaboral.idemp = pla_recmarcacion.idemp " _
        + vbCr + "WHERE (((pla_periodolaboral.fchini)<=CDate('" & FECH_ & "')) AND ((pla_periodolaboral.fchfin)>CDate('" & FECH_ & "') Or (pla_periodolaboral.fchfin) Is Null) AND ((pla_recmarcacion.dia)=CDate('" & FECH_ & "'))) " _
        + vbCr + "GROUP BY pla_periodolaboral.idemp, pla_empleados.idarea, pla_empleados.idcargo, pla_empleados.nombre"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    Set hallarPersonal = xRs
End Function

Private Function llenarDefinirGlobales() As Boolean
    Dim FCHINI_ As String
    Dim FCHFIN_ As String
    
    FCHINI_ = TxtFchAsisIni.Valor
    FCHFIN_ = TxtFchAsisFin.Valor
    
    ' CONSULTA DE INGRESANTES
    cSQL = "SELECT pla_periodolaboral.idemp, pla_empleados.numdoc, pla_empleados.nombre, pla_empleados.idarea, mae_area.descripcion AS desarea, pla_empleados.idcargo, mae_cargo.descripcion AS descargo, pla_periodolaboral.fchini " _
        + vbCr + "FROM ((pla_periodolaboral LEFT JOIN pla_empleados ON pla_periodolaboral.idemp = pla_empleados.id) LEFT JOIN mae_area ON pla_empleados.idarea = mae_area.id) LEFT JOIN mae_cargo ON pla_empleados.idcargo = mae_cargo.id " _
        + vbCr + "WHERE (((pla_periodolaboral.fchini)>=CDate('" & FCHINI_ & "') And (pla_periodolaboral.fchini)<=CDate('" & FCHFIN_ & "')));"
    
    Set xRsIng = Nothing
    RST_Busq xRsIng, cSQL, xCon
    
    ' CONSULTA DE SALIENTES
    cSQL = "SELECT pla_periodolaboral.idemp, pla_empleados.numdoc, pla_empleados.nombre, pla_empleados.idarea, mae_area.descripcion AS desarea, pla_empleados.idcargo, mae_cargo.descripcion AS descargo, pla_periodolaboral.fchfin " _
        + vbCr + "FROM ((pla_periodolaboral LEFT JOIN pla_empleados ON pla_periodolaboral.idemp=pla_empleados.id) LEFT JOIN mae_area ON pla_empleados.idarea=mae_area.id) LEFT JOIN mae_cargo ON pla_empleados.idcargo=mae_cargo.id " _
        + vbCr + "WHERE (((pla_periodolaboral.fchfin)>=CDate('" & FCHINI_ & "') And (pla_periodolaboral.fchfin)<=CDate('" & FCHFIN_ & "')));"
    
    Set xRsSal = Nothing
    RST_Busq xRsSal, cSQL, xCon
    
    ' CONSULTA DE RE-INGRESANTES
    cSQL = "SELECT pla_periodolaboral.corr, pla_periodolaboral.idemp, pla_empleados.idarea, pla_empleados.idcargo, pla_periodolaboral.fchini, pla_empleados.nombre " _
        + vbCr + "FROM pla_periodolaboral LEFT JOIN pla_empleados ON pla_periodolaboral.idemp = pla_empleados.id " _
        + vbCr + "WHERE (((pla_periodolaboral.corr)>1) AND ((pla_periodolaboral.fchini)>=CDate('" & FCHINI_ & "') And (pla_periodolaboral.fchini)<=CDate('" & FCHFIN_ & "')));"
    
    Set xRsRei = Nothing
    RST_Busq xRsRei, cSQL, xCon
    
    ' SE HALLA EL TOTAL DE PERSONAS EN LA FECHAS INDICADAS
    Set xRsPers = hallarPersonal(FCHINI_)
    
    ' SE HALLAN LOS CRITERIOS DE AGRUPAMIENTO
    If OptAgrupar(0).Value Then ' SEGUN AREA
        cSQL = "SELECT * FROM mae_area"
    ElseIf OptAgrupar(1).Value Then ' SEGUN CARGO
        cSQL = "SELECT * FROM mae_cargo"
    End If
    
    Set xRsCriAgr = Nothing
    RST_Busq xRsCriAgr, cSQL, xCon
    
    If xRsCriAgr.RecordCount = 0 Then llenarDefinirGlobales = False: Exit Function
    If xRsCriAgr.State = 0 Then llenarDefinirGlobales = False: Exit Function
    
    llenarDefinirGlobales = True
End Function

Private Sub mostrarDetalle(ID_ As Double, FCH_ As String)
    Dim FILTRO_ As String
    Dim FILAS_ As Integer
    Dim A As Integer
    
    If OptAgrupar(0).Value Then ' SEGUN AREA
        FILTRO_ = "idarea = "
    ElseIf OptAgrupar(1).Value Then ' SEGUN CARGO
        FILTRO_ = "idcargo = "
    End If
    
    fg(2).Rows = 1
    FILAS_ = 0
    
    lblDia.Caption = Format(FCH_, FORMAT_DATE)
    
    xRsCriAgr.Filter = "id = " & ID_
    lblDetalle.Caption = NulosC(xRsCriAgr("descripcion"))
    
    ' PERSONAL
    Set xRsPers = Nothing
    Set xRsPers = hallarPersonal(FCH_)
    xRsPers.Filter = adFilterNone
    xRsPers.Filter = FILTRO_ & ID_
    If xRsPers.RecordCount > 0 Then xRsPers.MoveFirst
    For A = 1 To xRsPers.RecordCount
        fg(2).Rows = fg(2).Rows + 1
        FILAS_ = FILAS_ + 1
        fg(2).TextMatrix(A, 2) = NulosC(xRsPers("nombre"))
        xRsPers.MoveNext
    Next A
    
    ' INGRESANTES
    xRsIng.Filter = adFilterNone
    xRsIng.Filter = FILTRO_ & ID_ & " And fchini = " & FCH_
    If xRsIng.RecordCount > 0 Then xRsIng.MoveFirst
    If xRsIng.RecordCount > FILAS_ Then FILAS_ = xRsIng.RecordCount: fg(2).Rows = FILAS_ + 1
    For A = 1 To xRsIng.RecordCount
        fg(2).TextMatrix(A, 3) = NulosC(xRsIng("nombre"))
        xRsIng.MoveNext
    Next A
    
    ' SALIENTES
    xRsSal.Filter = adFilterNone
    xRsSal.Filter = FILTRO_ & ID_ & " And fchfin = " & FCH_
    If xRsSal.RecordCount > 0 Then xRsSal.MoveFirst
    If xRsSal.RecordCount > FILAS_ Then FILAS_ = xRsSal.RecordCount: fg(2).Rows = FILAS_ + 1
    For A = 1 To xRsSal.RecordCount
        fg(2).TextMatrix(A, 4) = NulosC(xRsSal("nombre"))
        xRsSal.MoveNext
    Next A
    
    ' RE-INGRESANTES
    xRsRei.Filter = adFilterNone
    xRsRei.Filter = FILTRO_ & ID_ & " And fchini = " & FCH_
    If xRsRei.RecordCount > 0 Then xRsRei.MoveFirst
    If xRsRei.RecordCount > FILAS_ Then FILAS_ = xRsRei.RecordCount: fg(2).Rows = FILAS_ + 1
    For A = 1 To xRsRei.RecordCount
        fg(2).TextMatrix(A, 5) = NulosC(xRsRei("nombre"))
        xRsRei.MoveNext
    Next A
    
    ' Se llena el orden
    For A = 1 To fg(2).Rows - 1
        fg(2).TextMatrix(A, 1) = A
    Next A
End Sub

Private Sub hallarConsulta()
    Dim A As Integer
    Dim LIMITE_ As Integer
    Dim FILTRO_ As String
    Dim D_ As Date
    Dim COLUMNA_ As Integer
    Dim FILA_ As Integer
            
    If (Not IsDate(TxtFchAsisIni.Valor)) Or (Not IsDate(TxtFchAsisFin.Valor)) Then Exit Sub
    
    CentrarFrm FraProgreso
    FraProgreso.Visible = True
    lblproc.Caption = "Procesando los datos"
    
    If Not llenarDefinirGlobales Then Exit Sub
    
    ' SE LLENA EL RESUMEN
    FraProgreso.Refresh
    lblproc.Caption = "Agrupando la información"
    
    fg(1).Rows = 1
    
    With fg(1)
        ' Se llena el criterio de agrupamiento
        xRsCriAgr.MoveFirst
        For A = 1 To xRsCriAgr.RecordCount
            .Rows = .Rows + 1
            .TextMatrix(A, 1) = NulosN(xRsCriAgr("id"))
            .TextMatrix(A, 2) = NulosC(xRsCriAgr("descripcion"))
            xRsCriAgr.MoveNext
        Next A
    End With

    ' SE LLENA EL DETALLE
    fg(0).Rows = 1
    fg(0).Cols = 4
        
    ' Se agrega columna para los totales
    fg(0).Cols = fg(0).Cols + 1
    
    With fg(0)
        ' Se llena el criterio de agrupamiento
        xRsCriAgr.MoveFirst
        For A = 1 To xRsCriAgr.RecordCount
            .Rows = .Rows + 1
            .TextMatrix(A, 1) = NulosN(xRsCriAgr("id"))
            .TextMatrix(A, 2) = NulosC(xRsCriAgr("descripcion"))
            xRsCriAgr.MoveNext
        Next A
        
        LIMITE_ = .Rows - 1
        FILA_ = 1
        
        FraProgreso.Refresh
        lblproc.Caption = "Creando Reporte"
        PgBar.Min = 0
        PgBar.Max = LIMITE_
    
        For A = 1 To LIMITE_
            COLUMNA_ = 4
            FraProgreso.Refresh
            PgBar.Value = A
            
            For D_ = CDate(TxtFchAsisIni.Valor) To CDate(TxtFchAsisFin.Valor)
                If Weekday(D_) = 1 Then GoTo SIGUIENTE_
                
                If A = 1 Then
                    fg(0).Cols = fg(0).Cols + 1
                End If
                
                If COLUMNA_ = 4 Then
                    .AddItem "", FILA_ + 1 ' Ingresantes
                    .AddItem "", FILA_ + 1 ' Salientes
                    .AddItem "", FILA_ + 1 ' Re-Ingresantes
                    
                    .TextMatrix(FILA_, COLUMNA_ - 1) = "Num. Per."
                    .TextMatrix(FILA_ + 1, COLUMNA_ - 1) = "Num. Ing."
                    .TextMatrix(FILA_ + 2, COLUMNA_ - 1) = "Num. Sal."
                    .TextMatrix(FILA_ + 3, COLUMNA_ - 1) = "Num. Re-ing."
                End If
                            
                If OptAgrupar(0).Value Then ' SEGUN AREA
                    FILTRO_ = "idarea = "
                ElseIf OptAgrupar(1).Value Then ' SEGUN CARGO
                    FILTRO_ = "idcargo = "
                End If
                
                ' NUMERO DE PERSONAL
                Set xRsPers = Nothing
                Set xRsPers = hallarPersonal(NulosC(D_))
                xRsPers.Filter = adFilterNone
                xRsPers.Filter = FILTRO_ & NulosN(.TextMatrix(FILA_, 1))
                .TextMatrix(FILA_, COLUMNA_) = xRsPers.RecordCount
                                
                ' NUMERO DE INGRESANTES
                xRsIng.Filter = adFilterNone
                xRsIng.Filter = FILTRO_ & NulosN(.TextMatrix(FILA_, 1)) & " And fchini = " & NulosC(D_)
                .TextMatrix(FILA_ + 1, COLUMNA_) = xRsIng.RecordCount
                
                ' NUMERO DE SALIENTES
                xRsSal.Filter = adFilterNone
                xRsSal.Filter = FILTRO_ & NulosN(.TextMatrix(FILA_, 1)) & " And fchfin = " & NulosC(D_)
                .TextMatrix(FILA_ + 2, COLUMNA_) = xRsSal.RecordCount
                
                ' NUMERO DE RE-INGRESANTES
                xRsRei.Filter = adFilterNone
                xRsRei.Filter = FILTRO_ & NulosN(.TextMatrix(FILA_, 1)) & " And fchini = " & NulosC(D_)
                .TextMatrix(FILA_ + 3, COLUMNA_) = xRsRei.RecordCount
                
                COLUMNA_ = COLUMNA_ + 1
SIGUIENTE_:
            Next D_
            
            hallarTotales FILA_
            FILA_ = FILA_ + 4
        Next A
        
    End With
    
    FraProgreso.Refresh
    lblproc.Caption = "Aplicando cambios"
    configurarGrid
    GRID_AGRUPAR fg(0), 2
    GRID_AGRUPAR fg(1), 1
    
    FraProgreso.Visible = False
End Sub

Private Sub hallarTotales(FILA_ As Integer)
    Dim B As Integer
    Dim DIFERENCIAHORAS_ As Date
    Dim DIFERENCIAHORASCADENA_ As String
    Dim COLUMNA_ As Integer
    
    Dim NUMTOTALPER_ As Double
    Dim NUMTOTALING_ As Double
    Dim NUMTOTALSAL_ As Double
    Dim NUMTOTALREI_ As Double
    
    Dim NUMDIAS_ As Integer
    Dim CONTADOR_ As Integer
    
    NUMTOTALING_ = 0
    NUMTOTALPER_ = 0
    NUMTOTALSAL_ = 0
    NUMTOTALREI_ = 0
    NUMDIAS_ = 0
    CONTADOR_ = 4
    
    For CONTADOR_ = 4 To fg(0).Cols - 2
        NUMTOTALPER_ = NUMTOTALPER_ + fg(0).TextMatrix(FILA_, CONTADOR_)
        ' Total Horas
        NUMTOTALING_ = NUMTOTALING_ + fg(0).TextMatrix(FILA_ + 1, CONTADOR_)
        ' Total Horas Extra
        NUMTOTALSAL_ = NUMTOTALSAL_ + fg(0).TextMatrix(FILA_ + 2, CONTADOR_)
        ' Total Horas Tardanza
        NUMTOTALREI_ = NUMTOTALREI_ + fg(0).TextMatrix(FILA_ + 3, CONTADOR_)
        
        NUMDIAS_ = NUMDIAS_ + 1
    Next CONTADOR_
    
    fg(0).TextMatrix(FILA_, fg(0).Cols - 1) = NUMTOTALPER_ / NUMDIAS_
    fg(0).TextMatrix(FILA_, fg(0).Cols - 1) = Format(fg(0).TextMatrix(FILA_, fg(0).Cols - 1), "0.00")
    fg(0).TextMatrix(FILA_ + 1, fg(0).Cols - 1) = NUMTOTALING_
    fg(0).TextMatrix(FILA_ + 2, fg(0).Cols - 1) = NUMTOTALSAL_
    fg(0).TextMatrix(FILA_ + 3, fg(0).Cols - 1) = NUMTOTALREI_
        
    llenarDatosResumido NulosN(fg(0).TextMatrix(FILA_, 1)), (NUMTOTALPER_ / NUMDIAS_), NUMTOTALING_, NUMTOTALSAL_, NUMTOTALREI_
    
    ' Se agrupan los datos restantes
    GRID_COMBINAR fg(0), NulosN(FILA_), 2, NulosN(FILA_) + 3, 2, fg(0).TextMatrix(FILA_, 2), flexAlignLeftCenter, False, flexMergeFree, &H0&
    
    CALCULANDO_ = False
End Sub

Private Sub llenarDatosResumido(ID_ As Double, NUMPROMPER_ As Double, NUMTOTALING_ As Double, NUMTOTALSAL_ As Double, NUMTOTALREI_ As Double)
    Dim A As Integer
    
    For A = 1 To fg(1).Rows - 1
        If fg(1).TextMatrix(A, 1) = ID_ Then
            fg(1).TextMatrix(A, 3) = NUMPROMPER_
            fg(1).TextMatrix(A, 3) = Format(fg(1).TextMatrix(A, 3), "0.00")
            fg(1).TextMatrix(A, 4) = NUMTOTALING_
            fg(1).TextMatrix(A, 5) = NUMTOTALSAL_
            fg(1).TextMatrix(A, 6) = NUMTOTALREI_
            
            If NUMPROMPER_ = 0 Then
                fg(1).TextMatrix(A, 7) = 0
            Else
                fg(1).TextMatrix(A, 7) = ((NUMTOTALING_ - NUMTOTALSAL_) / NUMPROMPER_) * 100
            End If
            
            fg(1).TextMatrix(A, 7) = Format(fg(1).TextMatrix(A, 7), "0.00") & " %"
            Exit Sub
        End If
    Next A
End Sub

Private Sub configurarGrid()
    Dim FECHA_ As Date
    Dim NOMBRE_ As String
    Dim A As Integer
    Dim COLUMNA_ As Integer
    Dim D_ As Date
    
    FECHA_ = CDate(TxtFchAsisIni.Valor)
    NOMBRE_ = Format(FECHA_, FORMAT_DATE)
    COLUMNA_ = 4
    For D_ = CDate(TxtFchAsisIni.Valor) To CDate(TxtFchAsisFin.Valor)
        If Weekday(D_) = 1 Then GoTo SIGUIENTE_
        fg(0).TextMatrix(0, COLUMNA_) = Format(D_, FORMAT_DATE)
        COLUMNA_ = COLUMNA_ + 1
SIGUIENTE_:
    Next
    
    fg(0).TextMatrix(0, fg(0).Cols - 1) = "Totales"
End Sub

Private Sub fg_DblClick(Index As Integer)
    Dim FILA_ As Integer
    Dim RESIDUO_ As Integer
    
    Select Case Index
        Case 0
            RESIDUO_ = (fg(Index).Row) Mod 4
            Select Case RESIDUO_
                Case 0
                    FILA_ = fg(Index).Row - 3
                Case 1
                    FILA_ = fg(Index).Row
                Case 2
                    FILA_ = fg(Index).Row - 1
                Case 3
                    FILA_ = fg(Index).Row - 2
            End Select
            
            CentrarFrm frm(0)
            frm(0).Visible = True
            
            mostrarDetalle fg(Index).TextMatrix(FILA_, 1), fg(Index).TextMatrix(0, fg(0).Col)
    End Select
End Sub

Private Sub Form_Load()
    Quehace = 3
    'CONECTAR
    iniciarCampos
End Sub

Private Sub iniciarCampos()
    fg(0).AllowUserResizing = flexResizeColumns
    fg(0).Rows = 1
    fg(0).FrozenCols = 3
    fg(0).ColWidth(1) = 0
    
    fg(1).AllowUserResizing = flexResizeColumns
    fg(1).ExplorerBar = flexExSortShow
    fg(1).Rows = 1
    fg(1).ColWidth(1) = 0
    
    fg(2).FrozenCols = 1
    fg(2).Editable = flexEDNone
    
    OptAgrupar(0).Value = True
    
    TxtFchAsisIni.Valor = Date
    TxtFchAsisFin.Valor = Date
End Sub

Sub ExportarExcel(INDICE_ As Double)
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    Dim TITULO_ As String
    
    TITULO_ = "REPORTE DE HORAS DEL PERSONAL"

    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, fg(INDICE_), TITULO_, "", "", TITULO_
    Set X_EXPORT = Nothing
    MousePointer = vbDefault
    Exit Sub
error:
    MousePointer = vbDefault
    SHOW_ERROR Name, "Exportar"
End Sub

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub
    
    If Me.Width <= 12000 Then Me.Width = 12000
    If Me.Height <= 8100 Then Me.Height = 8100
    
    ' Se dimensiona la cabecera
    Frame2.Width = Me.Width - 130
    
    ' Se dimensiona el contenido
    Frame2.Width = Me.Width - 120
    Frame2.Height = Me.Height - 675
    
    TabOne1.Height = Frame2.Height - 810
    TabOne1.Width = Frame2.Width - 100
End Sub

Private Sub PbCerrar_Click(Index As Integer)
    frm(0).Visible = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 10 Then
        hallarConsulta
        'configurarGrid
    End If
    
    If Button.Index = 14 Then
        If TabOne1.CurrTab = 0 Then
            ExportarExcel 1
        Else
            ExportarExcel 0
        End If
    End If
    
    If Button.Index = 17 Then
        Set xCon = Nothing
        Unload Me
    End If
End Sub

Private Sub frm_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    frm(Index).ZOrder 0
End Sub

Private Sub frm_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With frm(Index)
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub
