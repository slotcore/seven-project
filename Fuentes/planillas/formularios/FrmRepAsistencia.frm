VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmRepAsistencia 
   Caption         =   "Planilla - Reporte de Asistencia"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   11820
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
            Picture         =   "FrmRepAsistencia.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepAsistencia.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepAsistencia.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepAsistencia.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepAsistencia.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepAsistencia.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepAsistencia.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepAsistencia.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepAsistencia.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepAsistencia.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepAsistencia.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepAsistencia.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepAsistencia.frx":277E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   609
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
      Begin VB.Frame Frame3 
         Caption         =   "Fecha de Asistencia"
         Height          =   975
         Left            =   60
         TabIndex        =   8
         Top             =   90
         Width           =   1965
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchAsisIni 
            Height          =   300
            Left            =   660
            TabIndex        =   9
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
            Left            =   660
            TabIndex        =   10
            Top             =   585
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
            TabIndex        =   12
            Top             =   285
            Width           =   375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fin"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   645
            Width           =   210
         End
      End
      Begin SizerOneLibCtl.TabOne TabOne1 
         Height          =   6360
         Left            =   30
         TabIndex        =   2
         Top             =   1065
         Width           =   11745
         _cx             =   20717
         _cy             =   11218
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
         CurrTab         =   1
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
            Height          =   6330
            Index           =   0
            Left            =   -12015
            TabIndex        =   3
            Top             =   15
            Width           =   11400
            _cx             =   20108
            _cy             =   11165
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
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmRepAsistencia.frx":2B10
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
            Height          =   6330
            Index           =   1
            Left            =   330
            TabIndex        =   4
            Top             =   15
            Width           =   11400
            _cx             =   20108
            _cy             =   11165
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
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmRepAsistencia.frx":2C0B
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
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   870
         Index           =   2
         Left            =   2070
         TabIndex        =   5
         Top             =   180
         Width           =   3030
         _cx             =   5345
         _cy             =   1535
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmRepAsistencia.frx":2D66
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
         Height          =   870
         Index           =   3
         Left            =   5160
         TabIndex        =   6
         Top             =   180
         Width           =   3030
         _cx             =   5345
         _cy             =   1535
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmRepAsistencia.frx":2DC5
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
         Height          =   870
         Index           =   4
         Left            =   8250
         TabIndex        =   7
         Top             =   180
         Width           =   3030
         _cx             =   5345
         _cy             =   1535
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmRepAsistencia.frx":2E23
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
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu00 
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "FrmRepAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rst As New ADODB.Recordset
Dim RstOrd As New ADODB.Recordset
Dim QueHace As Integer
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
'Dim xCon As New ADODB.Connection
Dim xTitulo As String
Dim CALCULANDO_ As Boolean

Private Sub hallarConsulta()
    ' El recordset para acceder a los datos
    Dim rs As ADODB.Recordset
    Dim A As Integer
    Dim B As Integer
    Dim contador As Integer
    Dim CAMPO_ As Date
    Dim CONTADOR_ As Integer
    ' Datos para la consulta
    Dim DNI As String
    Dim CARGO As String
    Dim EMPRESA As String
    Dim NUEVO_ As Boolean
    Dim IDEMP_ As Double
    
    Set rs = New ADODB.Recordset
    
    cPERSONAL = GENERAR_SQL_ID(fg(2), 1, " AND CONSULTASISTENCIA.[idemp]", "IN", True) ' Personal
    CCARGO = GENERAR_SQL_ID(fg(3), 1, " AND CONSULTASISTENCIA.idcargo", "IN", True) ' Cargo
    CAREA = GENERAR_SQL_ID(fg(4), 1, " AND CONSULTASISTENCIA.[idarea]", "IN", True) ' Area
    
    ' CONSULTA DE HORA DE ENTRADA
    cSQL = "TRANSFORM Min(CONSULTASISTENCIA.[hora]) AS MínDehora " _
        + vbCr + "SELECT CONSULTASISTENCIA.[idemp], CONSULTASISTENCIA.[nombre], CONSULTASISTENCIA.[idarea], CONSULTASISTENCIA.desarea, CONSULTASISTENCIA.idcargo, CONSULTASISTENCIA.descargo, CONSULTASISTENCIA.numdoc, CONSULTASISTENCIA.fching " _
        + vbCr + "FROM ( " _
        + vbCr + "SELECT pla_recmarcacion.idemp, pla_empleados.nombre, pla_empleados.idarea, mae_area.descripcion AS desarea, pla_empleados.idcargo, mae_cargo.descripcion AS descargo, pla_recmarcacion.dia, pla_recmarcacion.hora, pla_empleados.numdoc, pla_empleados.fching " _
        + vbCr + "FROM (((pla_recmarcacion INNER JOIN pla_empleados ON pla_recmarcacion.idemp = pla_empleados.id) LEFT JOIN mae_area ON pla_empleados.idarea = mae_area.id) LEFT JOIN mae_cargo ON pla_empleados.idcargo = mae_cargo.id) INNER JOIN mae_horarioemp ON pla_empleados.id = mae_horarioemp.idemp " _
        + vbCr + ") As CONSULTASISTENCIA " _
        + vbCr + "WHERE (((CONSULTASISTENCIA.dia)>=CDate('" & TxtFchAsisIni.Valor & "') And (CONSULTASISTENCIA.dia)<=CDate('" & TxtFchAsisFin.Valor & "'))) " & cPERSONAL & CCARGO & CAREA _
        + vbCr + "GROUP BY CONSULTASISTENCIA.[idemp], CONSULTASISTENCIA.[nombre], CONSULTASISTENCIA.[idarea], CONSULTASISTENCIA.desarea, CONSULTASISTENCIA.idcargo, CONSULTASISTENCIA.descargo, CONSULTASISTENCIA.numdoc, CONSULTASISTENCIA.fching " _
        + vbCr + "PIVOT CONSULTASISTENCIA.dia;"
    
    RST_Busq rs, cSQL, xCon
    
    fg(0).Rows = 1
    fg(1).Rows = 1
    
    ' Asignar el recordset al FlexGrid
    If rs.State = 0 Then Exit Sub
    If rs.RecordCount = 0 Then Exit Sub
                
    rs.MoveFirst
    fg(0).Cols = (2 * ((CDate(TxtFchAsisFin.Valor) - CDate(TxtFchAsisIni.Valor)) + 1)) + 9
    
    '********************************
    ' Se agrega columnas para :
    ' total dias
    ' total horas
    fg(0).Cols = fg(0).Cols + 3
    '********************************
    With rs
        For A = 1 To rs.RecordCount
            ' Detallado
            fg(0).Rows = fg(0).Rows + 1
            fg(0).TextMatrix(fg(0).Rows - 1, 1) = NulosN(.Fields("idemp"))
            fg(0).TextMatrix(fg(0).Rows - 1, 2) = NulosN(.Fields("idcargo"))
            fg(0).TextMatrix(fg(0).Rows - 1, 3) = NulosN(.Fields("idarea"))
            fg(0).TextMatrix(fg(0).Rows - 1, 4) = NulosC(.Fields("nombre"))
            fg(0).TextMatrix(fg(0).Rows - 1, 5) = NulosC(.Fields("numdoc"))
            fg(0).TextMatrix(fg(0).Rows - 1, 6) = NulosC(.Fields("descargo"))
            fg(0).TextMatrix(fg(0).Rows - 1, 7) = NulosC(.Fields("desarea"))
            fg(0).TextMatrix(fg(0).Rows - 1, 8) = Format(.Fields("fching"), FORMAT_DATE)
            
            CAMPO_ = CDate(TxtFchAsisIni.Valor)
            CONTADOR_ = 9
            For B = 1 To (CDate(TxtFchAsisFin.Valor) - CDate(TxtFchAsisIni.Valor) + 1)
                On Error Resume Next
                fg(0).TextMatrix(fg(0).Rows - 1, CONTADOR_) = Format(.Fields(Format(CAMPO_, "dd/mm/yyyy")), FORMAT_HORA_SIN_SEGUNDO)
                CONTADOR_ = CONTADOR_ + 2
                CAMPO_ = CAMPO_ + 1
            Next
            
            ' Resumido
            fg(1).Rows = fg(1).Rows + 1
            fg(1).TextMatrix(fg(1).Rows - 1, 1) = NulosN(.Fields("idemp"))
            fg(1).TextMatrix(fg(1).Rows - 1, 2) = NulosN(.Fields("idcargo"))
            fg(1).TextMatrix(fg(1).Rows - 1, 3) = NulosN(.Fields("idarea"))
            fg(1).TextMatrix(fg(1).Rows - 1, 4) = NulosC(.Fields("nombre"))
            fg(1).TextMatrix(fg(1).Rows - 1, 5) = NulosC(.Fields("numdoc"))
            fg(1).TextMatrix(fg(1).Rows - 1, 6) = NulosC(.Fields("descargo"))
            fg(1).TextMatrix(fg(1).Rows - 1, 7) = NulosC(.Fields("desarea"))
            fg(1).TextMatrix(fg(1).Rows - 1, 8) = Format(.Fields("fching"), FORMAT_DATE)
            
            .MoveNext
            If .EOF Then Exit For
        Next A
    End With

    'CONSULTA DE HORA DE SALIDA
    cSQL = "TRANSFORM Max(CONSULTASISTENCIA.[hora]) AS MáxDehora " _
        + vbCr + "SELECT CONSULTASISTENCIA.[idemp], CONSULTASISTENCIA.[nombre], CONSULTASISTENCIA.[idarea], CONSULTASISTENCIA.desarea, CONSULTASISTENCIA.idcargo, CONSULTASISTENCIA.descargo " _
        + vbCr + "FROM ( " _
        + vbCr + "SELECT pla_recmarcacion.idemp, pla_empleados.nombre, pla_empleados.idarea, mae_area.descripcion AS desarea, pla_empleados.idcargo, mae_cargo.descripcion AS descargo, pla_recmarcacion.dia, pla_recmarcacion.hora " _
        + vbCr + "FROM ((pla_recmarcacion LEFT JOIN pla_empleados ON pla_recmarcacion.idemp = pla_empleados.id) LEFT JOIN mae_area ON pla_empleados.idarea = mae_area.id) LEFT JOIN mae_cargo ON pla_empleados.idcargo = mae_cargo.id " _
        + vbCr + ") As CONSULTASISTENCIA " _
        + vbCr + "WHERE (((CONSULTASISTENCIA.dia)>=CDate('" & TxtFchAsisIni.Valor & "') And (CONSULTASISTENCIA.dia)<=CDate('" & TxtFchAsisFin.Valor & "'))) " & cPERSONAL & CCARGO & CAREA _
        + vbCr + "GROUP BY CONSULTASISTENCIA.[idemp], CONSULTASISTENCIA.[nombre], CONSULTASISTENCIA.[idarea], CONSULTASISTENCIA.desarea, CONSULTASISTENCIA.idcargo, CONSULTASISTENCIA.descargo " _
        + vbCr + "PIVOT CONSULTASISTENCIA.dia;"

    RST_Busq rs, cSQL, xCon
    
    With fg(0)
        For A = 1 To .Rows - 1
            CAMPO_ = CDate(TxtFchAsisIni.Valor)
            CONTADOR_ = 10
            rs.Filter = "idemp = " & NulosC(.TextMatrix(A, 1))
            If rs.RecordCount = 0 Then GoTo SIGUIENTE
            
            For B = 1 To (CDate(TxtFchAsisFin.Valor) - CDate(TxtFchAsisIni.Valor) + 1)
                On Error Resume Next
                .TextMatrix(A, CONTADOR_) = Format(rs.Fields(Format(CAMPO_, "dd/mm/yyyy")), FORMAT_HORA_SIN_SEGUNDO)
                
                ' si son horas iguales y no nulas
                If ((.TextMatrix(A, CONTADOR_) = .TextMatrix(A, CONTADOR_ - 1)) And (.TextMatrix(A, CONTADOR_) <> "")) Then
                    .Select A, CONTADOR_ - 1, A, CONTADOR_
                    .FillStyle = flexFillRepeat
                    .CellBackColor = &HB9B9FF
                End If
                
                CONTADOR_ = CONTADOR_ + 2
                CAMPO_ = CAMPO_ + 1
            Next B
SIGUIENTE:
        Next A
    End With
    
    Set rs = Nothing
     
    Dim LIMITE_ As Double
    Dim FILA_ As Integer
     
    LIMITE_ = fg(0).Rows - 1
    FILA_ = 1
    For A = 1 To LIMITE_
        NUEVO_ = True
        CONTADOR_ = 9
        For B = 1 To (CDate(TxtFchAsisFin.Valor) - CDate(TxtFchAsisIni.Valor) + 1)
            IDEMP_ = NulosN(fg(0).TextMatrix(FILA_, 1))
            calcularDatos IDEMP_, FILA_, CONTADOR_, NUEVO_
            CONTADOR_ = CONTADOR_ + 2
        Next B
        
        calcularDiasHoras FILA_
        FILA_ = FILA_ + 4
    Next A
End Sub

Private Sub calcularDatos(IDEMP_ As Double, FILA_ As Integer, COLUMNA_ As Integer, ByRef AGREGAR_ As Boolean)
    Dim xRs As New ADODB.Recordset
    Dim HORAINI_ As Date
    Dim HORAFIN_ As Date
    Dim HORAINIHORARIO_ As Date
    Dim HORAFINHORARIO_ As Date
    Dim HORAINIDESCANSO_ As Date
    Dim HORAFINDESCANSO_ As Date
    Dim DESCANSO_ As Date
    
    ' Buscamos horas de trabajo
    cSQL = "SELECT mae_horariohora.hingreso, mae_horariohora.hsalida " _
        + vbCr + "FROM (mae_horario LEFT JOIN mae_horariohora ON mae_horario.id = mae_horariohora.idhor) RIGHT JOIN mae_horarioemp ON mae_horario.id = mae_horarioemp.idhor " _
        + vbCr + "WHERE (((mae_horariohora.idhora)=1) AND ((mae_horarioemp.idemp)=" & IDEMP_ & "));"

    RST_Busq xRs, cSQL, xCon

    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub

    HORAINIHORARIO_ = xRs("hingreso")
    HORAFINHORARIO_ = xRs("hsalida")

    ' Buscamos horas de descanso
    cSQL = "SELECT mae_horariohora.hingreso, mae_horariohora.hsalida " _
        + vbCr + "FROM (mae_horario LEFT JOIN mae_horariohora ON mae_horario.id = mae_horariohora.idhor) RIGHT JOIN mae_horarioemp ON mae_horario.id = mae_horarioemp.idhor " _
        + vbCr + "WHERE (((mae_horariohora.idhora)=14) AND ((mae_horarioemp.idemp)=" & IDEMP_ & "));"
        
    RST_Busq xRs, cSQL, xCon

    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    HORAINIDESCANSO_ = xRs("hingreso")
    HORAFINDESCANSO_ = xRs("hsalida")
    DESCANSO_ = CDate(xRs("hsalida")) - CDate(xRs("hingreso"))
    
    HORAINI_ = CDate(fg(0).TextMatrix(FILA_, COLUMNA_))
    HORAFIN_ = CDate(fg(0).TextMatrix(FILA_, COLUMNA_ + 1))
    
    If AGREGAR_ Then
        fg(0).AddItem "", FILA_ + 1
        fg(0).AddItem "", FILA_ + 1
        fg(0).AddItem "", FILA_ + 1
    End If
    
    AGREGAR_ = False
    CALCULANDO_ = True
    fg(0).TextMatrix(FILA_ + 1, COLUMNA_) = "Tot. Hrs."
    fg(0).TextMatrix(FILA_ + 2, COLUMNA_) = "Hrs. Ext."
    fg(0).TextMatrix(FILA_ + 3, COLUMNA_) = "Hrs. Tar."
    fg(0).Select FILA_ + 1, COLUMNA_, FILA_ + 3, COLUMNA_
    fg(0).FillStyle = flexFillRepeat
    fg(0).CellForeColor = &HC0&       '&H00000000&
    
    fg(0).Select FILA_ + 1, COLUMNA_, FILA_ + 3, COLUMNA_ + 1
    fg(0).FillStyle = flexFillRepeat
    fg(0).CellBackColor = &HC0FFFF             '&H00000000&
    
    If fg(0).TextMatrix(FILA_, COLUMNA_ + 1) = fg(0).TextMatrix(FILA_, COLUMNA_) Then
        fg(0).TextMatrix(FILA_ + 1, COLUMNA_ + 1) = "00:00"
        fg(0).TextMatrix(FILA_ + 2, COLUMNA_ + 1) = "00:00"
        fg(0).TextMatrix(FILA_ + 3, COLUMNA_ + 1) = "00:00"
    Else
        If contieneDescanso(HORAINI_, HORAFIN_, HORAINIDESCANSO_, HORAFINDESCANSO_) Then
            fg(0).TextMatrix(FILA_ + 1, COLUMNA_ + 1) = Format(HORAFIN_ - HORAINI_ - DESCANSO_, "HH:mm")
        Else
            fg(0).TextMatrix(FILA_ + 1, COLUMNA_ + 1) = Format(HORAFIN_ - HORAINI_, "HH:mm")
        End If
        
        If HORAFIN_ > HORAFINHORARIO_ Then
            fg(0).TextMatrix(FILA_ + 2, COLUMNA_ + 1) = Format(HORAFIN_ - HORAFINHORARIO_, "HH:mm")
        Else
            fg(0).TextMatrix(FILA_ + 2, COLUMNA_ + 1) = "00:00"
        End If
        
        If HORAINI_ > HORAINIHORARIO_ Then
            fg(0).TextMatrix(FILA_ + 3, COLUMNA_ + 1) = Format(HORAINIHORARIO_ - HORAINI_, "HH:mm")
        Else
            fg(0).TextMatrix(FILA_ + 3, COLUMNA_ + 1) = "00:00"
        End If
    End If
    CALCULANDO_ = False
End Sub

Private Function contieneDescanso(HORAINI_ As Date, HORAFIN_ As Date, HORAINIDESCANSO_ As Date, HORAFINDESCANSO_ As Date) As Boolean
    Dim VERIFICO_ As Boolean
    
    VERIFICO_ = False
    
    If HORAINIDESCANSO_ > HORAINI_ Then
        If HORAFIN_ > HORAFINDESCANSO_ Then
            VERIFICO_ = True
        End If
    End If
    
    contieneDescanso = VERIFICO_
End Function

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0 ' Consultar
            hallarConsulta
            configurarGrid
    End Select
End Sub

Private Sub calcularDiasHoras(FILA_ As Integer)
    Dim B As Integer
    Dim DIFERENCIAHORAS_ As Date
    Dim DIFERENCIAHORASCADENA_ As String
    Dim COLUMNA_ As Integer
    Dim NUMHORASTOTALNETA_ As Double
    Dim NUMHORASTOTAL_ As Double
    Dim NUMHORASEXTRA_ As Double
    Dim NUMHORASTARDANZA_ As Double
    Dim NUMDIAS_ As Integer
    Dim CONTADOR_ As Integer
    
    NUMHORASTOTAL_ = 0
    NUMDIAS_ = 0
    CONTADOR_ = 10
    For CONTADOR_ = 10 To fg(0).Cols - 3 Step 2
        If NulosC(fg(0).TextMatrix(FILA_, CONTADOR_)) <> "" Then
            If fg(0).TextMatrix(FILA_, CONTADOR_) = fg(0).TextMatrix(FILA_, CONTADOR_ - 1) Then
                DIFERENCIAHORASCADENA_ = "00:00"
                NUMDIAS_ = NUMDIAS_ + 1
            Else
                If Not IsDate(fg(0).TextMatrix(FILA_, CONTADOR_)) Or Not IsDate(fg(0).TextMatrix(FILA_, CONTADOR_ - 1)) Then
                    DIFERENCIAHORASCADENA_ = "00:00"
                Else
                    ' Total Horas Neta
                    DIFERENCIAHORAS_ = CDate(fg(0).TextMatrix(FILA_, CONTADOR_)) - CDate(fg(0).TextMatrix(FILA_, CONTADOR_ - 1))
                    DIFERENCIAHORASCADENA_ = Format(DIFERENCIAHORAS_, "HH:mm")
                    NUMHORASTOTALNETA_ = NUMHORASTOTALNETA_ + convertirHorasNumero(DIFERENCIAHORASCADENA_)
                    ' Total Horas
                    NUMHORASTOTAL_ = NUMHORASTOTAL_ + convertirHorasNumero(fg(0).TextMatrix(FILA_ + 1, CONTADOR_))
                    ' Total Horas Extra
                    NUMHORASEXTRA_ = NUMHORASEXTRA_ + convertirHorasNumero(fg(0).TextMatrix(FILA_ + 2, CONTADOR_))
                    ' Total Horas Tardanza
                    NUMHORASTARDANZA_ = NUMHORASTARDANZA_ + convertirHorasNumero(fg(0).TextMatrix(FILA_ + 3, CONTADOR_))
                    
                    NUMDIAS_ = NUMDIAS_ + 1
                End If
            End If
        End If
    Next CONTADOR_
    
    CALCULANDO_ = True
    fg(0).TextMatrix(FILA_, fg(0).Cols - 3) = NUMDIAS_
    
    fg(0).TextMatrix(FILA_, fg(0).Cols - 2) = "Hrs. Net."
    fg(0).TextMatrix(FILA_, fg(0).Cols - 1) = convertirNumeroHoras(NUMHORASTOTALNETA_)
    
    fg(0).TextMatrix(FILA_ + 1, fg(0).Cols - 2) = "Hrs."
    fg(0).TextMatrix(FILA_ + 2, fg(0).Cols - 2) = "Hrs. Ext."
    fg(0).TextMatrix(FILA_ + 3, fg(0).Cols - 2) = "Hrs. Tar."
    fg(0).TextMatrix(FILA_ + 1, fg(0).Cols - 1) = convertirNumeroHoras(NUMHORASTOTAL_)
    fg(0).TextMatrix(FILA_ + 2, fg(0).Cols - 1) = convertirNumeroHoras(NUMHORASEXTRA_)
    fg(0).TextMatrix(FILA_ + 3, fg(0).Cols - 1) = convertirNumeroHoras(NUMHORASTARDANZA_)
    
    fg(0).Select FILA_, fg(0).Cols - 2, FILA_ + 3, fg(0).Cols - 1
    fg(0).FillStyle = flexFillRepeat
    fg(0).CellBackColor = &HFFF0D7          '&H00000000&
    fg(0).Select FILA_, fg(0).Cols - 2, FILA_ + 3, fg(0).Cols - 2
    fg(0).FillStyle = flexFillRepeat
    fg(0).CellForeColor = &HFF&
    
    llenarDatosResumido NulosN(fg(0).TextMatrix(FILA_, 1)), NUMHORASTOTAL_, NUMHORASEXTRA_, NUMHORASTARDANZA_
    
    ' Se agrupan los datos restantes
    For B = 4 To 8
        GRID_COMBINAR fg(0), NulosN(FILA_), B, NulosN(FILA_) + 3, B, fg(0).TextMatrix(FILA_, B), flexAlignLeftCenter, False, flexMergeRestrictRows, &H0&
    Next
    
    GRID_COMBINAR fg(0), NulosN(FILA_), fg(0).Cols - 3, NulosN(FILA_) + 3, fg(0).Cols - 3, fg(0).TextMatrix(FILA_, fg(0).Cols - 3), flexAlignCenterCenter, False, flexMergeRestrictRows, &H0&
    
    CALCULANDO_ = False
End Sub

Private Sub llenarDatosResumido(IDEMP_ As Double, NUMHORASTOTAL_ As Double, NUMHORASEXTRA_ As Double, NUMHORASTARDANZA_ As Double)
    Dim A As Integer
    
    For A = 1 To fg(1).Rows - 1
        If fg(1).TextMatrix(A, 1) = IDEMP_ Then
           fg(1).TextMatrix(A, 9) = convertirNumeroHoras(NUMHORASTOTAL_)
           fg(1).TextMatrix(A, 10) = convertirNumeroHoras(NUMHORASEXTRA_)
           fg(1).TextMatrix(A, 11) = convertirNumeroHoras(NUMHORASTARDANZA_)
           Exit Sub
        End If
    Next A
End Sub

Private Function convertirHorasNumero(HORACADENA_ As String) As Double
    Dim h() As String
    Dim tiempo As Double
    h = Split(HORACADENA_, ":")
    tiempo = Val(h(0)) + (Val(h(1)) / 60)
    
    convertirHorasNumero = tiempo
End Function

Private Function convertirNumeroHoras(HORANUMERO_ As Double) As String
    Dim xHorEst As String
    
    xHorEst = Format(Int(HORANUMERO_), "00")
    xHorEst = xHorEst & ":" & Format(((HORANUMERO_ * 60) Mod 60), "00")
    
    convertirNumeroHoras = xHorEst
End Function

'Private Sub CmdBusCargo_Click()
'    Dim xform As New eps_librerias.FormBuscar
'    Dim xRs As New ADODB.Recordset
'    Dim xCampos(2, 4) As String
'
'    xCampos(0, 0) = "Cargo":         xCampos(0, 1) = "CARGO":           xCampos(0, 2) = "1500":     xCampos(0, 3) = "C"
'    xCampos(1, 0) = "Descripcion":   xCampos(1, 1) = "DESCRIPCION":     xCampos(1, 2) = "4000":     xCampos(1, 3) = "C"
'
'    cSQL = "SELECT CARGO, DESCRIPCION " _
'        + vbCr + "From TEMPUS_CARGOS "
'
'    xform.SQLCad = cSQL
'
'    xform.Titulo = "Buscando Cargos"
'    xform.FormaBusca = Principio
'    xform.Criterio = ""
'    xform.Ordenado = "CARGO"
'    xform.CampoBusca = "DESCRIPCION"
'    Set xform.Coneccion = xCon
'
'    ' Inicia tabla de busqueda
'    Set xRs = xform.BuscarReg(xCampos)
'
'    If xRs.State = 1 Then
'        TxtIdCargo.Text = xRs("CARGO")                 ' Descripcion del producto
'        LblNomCargo.Caption = xRs("DESCRIPCION")       ' Descripcion de la UM
'    End If
'End Sub
'
'Private Sub CmdBusEmpresa_Click()
'    Dim xform As New eps_librerias.FormBuscar
'    Dim xRs As New ADODB.Recordset
'    Dim xCampos(2, 4) As String
'
'
'    xCampos(0, 0) = "Empresa":       xCampos(0, 1) = "EMPRESA":    xCampos(0, 2) = "1500":     xCampos(0, 3) = "C"
'    xCampos(1, 0) = "Descripcion":   xCampos(1, 1) = "NOMBRE":     xCampos(1, 2) = "4000":     xCampos(1, 3) = "C"
'
'    cSQL = "SELECT EMPRESA, NOMBRE " _
'        + vbCr + "From TEMPUS_EMPRESAS "
'
'    xform.SQLCad = cSQL
'
'    xform.Titulo = "Buscando Cargos"
'    xform.FormaBusca = Principio
'    xform.Criterio = ""
'    xform.Ordenado = "EMPRESA"
'    xform.CampoBusca = "NOMBRE"
'    Set xform.Coneccion = xCon
'
'    'Inicia tabla de busqueda
'    Set xRs = xform.BuscarReg(xCampos)
'
'    If xRs.State = 1 Then
'        TxtIdEmpres.Text = xRs("EMPRESA")                 ' Descripcion del producto
'        LblNomEmpresa.Caption = xRs("NOMBRE")             ' Descripcion de la UM
'    End If
'End Sub

Private Sub configurarGrid()
    Dim FECHA_ As Date
    Dim NOMBRE_ As String
    Dim A As Integer
    Dim CONTADOR_ As Integer
    
    FECHA_ = CDate(TxtFchAsisIni.Valor)
    NOMBRE_ = Format(FECHA_, FORMAT_DATE)
    CONTADOR_ = 9
    For A = 1 To (CDate(TxtFchAsisFin.Valor) - CDate(TxtFchAsisIni.Valor)) + 1
        GRID_COMBINAR fg(0), 0, CONTADOR_, 0, CONTADOR_ + 1, NOMBRE_, flexAlignCenterCenter, True, flexMergeFixedOnly, , &H8000000F, False
        FECHA_ = FECHA_ + 1
        NOMBRE_ = Format(FECHA_, FORMAT_DATE)
        CONTADOR_ = CONTADOR_ + 2
    Next
    
    fg(0).TextMatrix(0, fg(0).Cols - 3) = "Dias Trab."
    GRID_COMBINAR fg(0), 0, fg(0).Cols - 2, 0, fg(0).Cols - 1, "Totales", flexAlignCenterCenter, True, flexMergeFree, , &H8000000F
End Sub

Private Sub fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim xCampos() As String
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String
    Dim nTitulo As String

    If Index = 2 Then
        ReDim xCampos(5, 4) As String
        
        xCampos(0, 0) = "DNI":                  xCampos(0, 1) = "numdoc":       xCampos(0, 2) = "1000":      xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
        xCampos(1, 0) = "Apellidos y Nombres":  xCampos(1, 1) = "nombre":       xCampos(1, 2) = "3100":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
        xCampos(2, 0) = "Fch. Ing.":            xCampos(2, 1) = "fching":       xCampos(2, 2) = "1000":     xCampos(2, 3) = "C":    xCampos(2, 4) = "C"
        xCampos(3, 0) = "Cargo":                xCampos(3, 1) = "cargo":        xCampos(3, 2) = "1500":     xCampos(3, 3) = "C":    xCampos(3, 4) = "C"
        xCampos(4, 0) = "Area":                 xCampos(4, 1) = "area":         xCampos(4, 2) = "1500":     xCampos(4, 3) = "C":    xCampos(4, 4) = "C"
                
        ' generar la lista de personal para no considerar en la lista
        nSQLId = GENERAR_SQL_ID(fg(2), 1, " AND pla_empleados.id", "NOT IN", True)
        
        ' generar la consulta
        cSQL = "SELECT pla_empleados.id AS idemp, pla_empleados.nombre, pla_empleados.numdoc, pla_empleados.fching, pla_empleados.fchcese, mae_area.descripcion AS area, mae_cargo.descripcion AS cargo " _
            + vbCr + "FROM (pla_empleados LEFT JOIN mae_area ON pla_empleados.idarea = mae_area.id) LEFT JOIN mae_cargo ON pla_empleados.idcargo = mae_cargo.id " _
            + vbCr + "WHERE (((pla_empleados.fchcese) Is Null)) " & nSQLId _
            + vbCr + "ORDER BY pla_empleados.nombre;"
            
        nTitulo = "Buscando Personal"
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
                      
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        fg(Index).Rows = fg(Index).Rows + 1
        'fg(Index).Select fg(Index).Rows - 1, 1
        
        fg(Index).TextMatrix(fg(Index).Row, 1) = NulosN(xRs("idemp"))
        fg(Index).TextMatrix(fg(Index).Row, 2) = NulosC(xRs("nombre"))
    End If
    
    If Index = 3 Then ' Cargo
        ReDim xCampos(2, 4) As String
        
        xCampos(0, 0) = "Id":           xCampos(0, 1) = "id":               xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
        xCampos(1, 0) = "Cargo":        xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "3100":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
                
        ' generar la lista de personal para no considerar en la lista
        nSQLId = GENERAR_SQL_ID(fg(3), 1, " AND mae_cargo.id", "NOT IN", True)
        
        ' generar la consulta
        cSQL = "SELECT mae_cargo.id, mae_cargo.descripcion " _
            + vbCr + "FROM mae_cargo " _
            + vbCr + "WHERE ((mae_cargo.id) Is Not null) " & nSQLId
            
        nTitulo = "Buscando Personal"
        
        Set xRs = Nothing
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
                      
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        fg(Index).Rows = fg(Index).Rows + 1
        'fg(Index).Select fg(Index).Rows - 1, 1
        
        fg(Index).TextMatrix(fg(Index).Row, 1) = NulosN(xRs("id"))
        fg(Index).TextMatrix(fg(Index).Row, 2) = NulosC(xRs("descripcion"))
    End If
    
    If Index = 4 Then ' Area
        ReDim xCampos(2, 4) As String
        
        xCampos(0, 0) = "Id":           xCampos(0, 1) = "id":               xCampos(0, 2) = "1000":      xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
        xCampos(1, 0) = "Area":         xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "3100":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
                
        ' generar la lista de personal para no considerar en la lista
        nSQLId = GENERAR_SQL_ID(fg(4), 1, " AND mae_area.id", "NOT IN", True)
        
        ' generar la consulta
        cSQL = "SELECT mae_area.id, mae_area.descripcion " _
            + vbCr + "FROM mae_area " _
            + vbCr + "WHERE ((mae_area.id) Is Not Null) " & nSQLId
            
        nTitulo = "Buscando Personal"
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
                      
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        fg(Index).Rows = fg(Index).Rows + 1
        'fg(Index).Select fg(Index).Rows - 1, 1
        
        fg(Index).TextMatrix(fg(Index).Row, 1) = NulosN(xRs("id"))
        fg(Index).TextMatrix(fg(Index).Row, 2) = NulosC(xRs("descripcion"))
    End If
End Sub

Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Index = 0 Then
        If Row = 0 Then Exit Sub
        If CALCULANDO_ Then Exit Sub
        If Col < 9 And Col > 0 Then
            Exit Sub
        Else
            ' Se da formato a la hora ingresada
            fg(0).TextMatrix(Row, Col) = Format(fg(0).TextMatrix(Row, Col), FORMAT_HORA_SIN_SEGUNDO)
            If (Col Mod 2 = 0) Then ' Hora de Salida
                fg(0).Select Row, Col - 1, Row, Col
                fg(0).FillStyle = flexFillRepeat
                fg(0).CellBackColor = &H80000005
            Else ' Hora de Entrada
                fg(0).Select Row, Col, Row, Col + 1
                fg(0).FillStyle = flexFillRepeat
                fg(0).CellBackColor = &H80000005
            End If
        End If
    End If
End Sub

Private Sub Fg_EnterCell(Index As Integer)
    If Index = 0 Then
        If fg(0).Row = 0 Then Exit Sub
        If fg(0).Col < 9 And fg(0).Col > 0 Then
            fg(0).Editable = flexEDNone
        Else
            fg(0).ColEditMask(fg(0).Col) = "##:##"
            fg(0).Editable = flexEDKbdMouse
        End If
    End If
End Sub

Private Sub fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 2 Then
        If Button = 2 Then
            PopupMenu menu
        End If
    End If
End Sub

Private Sub Form_Load()
    QueHace = 3
    'CONECTAR
    iniciarCampos
End Sub

Private Sub iniciarCampos()
    fg(0).AllowUserResizing = flexResizeColumns
    fg(0).ExplorerBar = flexExSortShow
    fg(0).Rows = 1
    fg(0).FrozenCols = 8
    fg(0).ColWidth(1) = 0
    fg(0).ColWidth(2) = 0
    fg(0).ColWidth(3) = 0
    
    fg(1).AllowUserResizing = flexResizeColumns
    fg(1).ExplorerBar = flexExSortShow
    fg(1).Rows = 1
    fg(1).FrozenCols = 8
    fg(1).ColWidth(1) = 0
    fg(1).ColWidth(2) = 0
    fg(1).ColWidth(3) = 0

    TxtFchAsisIni.Valor = Date
    TxtFchAsisFin.Valor = Date
    
    fg(2).ColWidth(1) = 0
    fg(2).ColComboList(2) = "|..."
    fg(2).Editable = flexEDKbdMouse
    fg(3).ColWidth(1) = 0
    fg(3).ColComboList(2) = "|..."
    fg(3).Editable = flexEDKbdMouse
    fg(4).ColWidth(1) = 0
    fg(4).ColComboList(2) = "|..."
    fg(4).Editable = flexEDKbdMouse
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
    
    TabOne1.Height = Frame2.Height - 1170
    TabOne1.Width = Frame2.Width - 100
End Sub

Private Sub menu00_Click()
    If fg(2).Row < fg(2).FixedRows Then Exit Sub
    fg(2).RemoveItem fg(2).Row
    
    If fg(2).Rows = fg(2).FixedRows Then
        fg(2).Rows = fg(2).Rows + 1
        fg(2).TextMatrix(1, 3) = "TODOS"
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 10 Then
        hallarConsulta
        configurarGrid
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
