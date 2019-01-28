VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRepLinea 
   Caption         =   "Produccion  -  Reporte de Lineas"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   13320
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   3570
      TabIndex        =   4
      Top             =   3810
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   225
         TabIndex        =   5
         Top             =   465
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Shape Shape1 
         Height          =   765
         Index           =   0
         Left            =   60
         Top             =   90
         Width           =   5805
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "lbl(1)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1395
         TabIndex        =   8
         Top             =   180
         Width           =   2670
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Procesando:"
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
         Index           =   0
         Left            =   225
         TabIndex        =   7
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Interrumpir = ESC"
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
         Index           =   2
         Left            =   4170
         TabIndex        =   6
         Top             =   180
         Width           =   1530
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   13320
      _ExtentX        =   23495
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   11070
         Top             =   0
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
               Picture         =   "FrmRepLinea.frx":0000
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepLinea.frx":0544
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepLinea.frx":08D6
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepLinea.frx":0A5A
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepLinea.frx":0EAE
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepLinea.frx":0FC6
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepLinea.frx":150A
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepLinea.frx":1A4E
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepLinea.frx":1B62
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepLinea.frx":1C76
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepLinea.frx":20CA
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepLinea.frx":2236
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepLinea.frx":277E
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepLinea.frx":2A98
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame6 
      BorderStyle     =   0  'None
      Caption         =   "Reporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   30
      TabIndex        =   0
      Top             =   330
      Width           =   13350
      Begin VB.Frame Frame2 
         Caption         =   "[ Tipo de Consulta ]"
         Height          =   885
         Left            =   11640
         TabIndex        =   17
         Top             =   0
         Width           =   1665
         Begin VB.OptionButton Opt 
            Caption         =   "Resumido"
            Height          =   285
            Index           =   0
            Left            =   150
            TabIndex        =   19
            Top             =   240
            Width           =   1035
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Detallado"
            Height          =   285
            Index           =   1
            Left            =   150
            TabIndex        =   18
            Top             =   510
            Width           =   1215
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "[Fech. Prod.]"
         Height          =   885
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1875
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchDesde 
            Height          =   300
            Left            =   555
            TabIndex        =   10
            Top             =   210
            Width           =   1275
            _ExtentX        =   2249
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchHasta 
            Height          =   300
            Left            =   555
            TabIndex        =   11
            Top             =   540
            Width           =   1275
            _ExtentX        =   2249
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
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            Height          =   195
            Left            =   45
            TabIndex        =   13
            Top             =   255
            Width           =   510
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Hasta:"
            Height          =   195
            Left            =   45
            TabIndex        =   12
            Top             =   585
            Width           =   465
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   6900
         Index           =   0
         Left            =   30
         TabIndex        =   1
         Top             =   930
         Width           =   13245
         _cx             =   23363
         _cy             =   12171
         _ConvInfo       =   1
         Appearance      =   1
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   22
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmRepLinea.frx":2E2A
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
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   795
         Index           =   3
         Left            =   1920
         TabIndex        =   14
         ToolTipText     =   "Buscar Linea"
         Top             =   70
         Width           =   3345
         _cx             =   5900
         _cy             =   1402
         _ConvInfo       =   1
         Appearance      =   1
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
         SelectionMode   =   1
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
         FormatString    =   $"FrmRepLinea.frx":308F
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   2
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   -1  'True
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
         Height          =   795
         Index           =   4
         Left            =   8460
         TabIndex        =   15
         ToolTipText     =   "Buscar Supervisor"
         Top             =   70
         Width           =   3135
         _cx             =   5521
         _cy             =   1402
         _ConvInfo       =   1
         Appearance      =   1
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
         SelectionMode   =   1
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
         FormatString    =   $"FrmRepLinea.frx":30EC
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   2
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   -1  'True
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
         Height          =   795
         Index           =   5
         Left            =   5300
         TabIndex        =   16
         ToolTipText     =   "Buscar Personal"
         Top             =   70
         Width           =   3135
         _cx             =   5530
         _cy             =   1402
         _ConvInfo       =   1
         Appearance      =   1
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
         SelectionMode   =   1
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
         FormatString    =   $"FrmRepLinea.frx":314E
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   2
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   -1  'True
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
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTO_01"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   0
      Width           =   1365
   End
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu00 
         Caption         =   "Insertar Item"
      End
      Begin VB.Menu separador 
         Caption         =   "-"
      End
      Begin VB.Menu menu01 
         Caption         =   "Eliminar Item"
      End
   End
End
Attribute VB_Name = "FrmRepLinea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CARGO As Boolean
Dim cSQL As String
Dim RstResumido As New ADODB.Recordset
Dim RstNumper As New ADODB.Recordset
Dim RstTareas As New ADODB.Recordset
Dim INDICE_ As Integer
Dim AGREGANDO_ As Boolean
Dim INTERRUMPIR_ As Boolean

Private Sub Buscar()
    ' Se verIfica si esta correcta la informacion
    If Not verificarDatos Then Exit Sub
    
    CARGO = True
    
    If Opt(0).Value Then fg(5).Rows = fg(5).FixedRows: fg(5).Rows = fg(5).Rows + 1
    
    generarConsulta Opt(0).Value, Opt(1).Value
    llenarDatos Opt(0).Value, Opt(1).Value
    hallarDatosRestantes Opt(0).Value, Opt(1).Value
    configurarGrid Opt(0).Value, Opt(1).Value
End Sub

Private Function verificarDatos() As Boolean
    Dim VERIFICO_ As Boolean
    Dim MENSAJE_ As String
    
    VERIFICO_ = True
    If (TxtFchDesde.valor = "" Or TxtFchHasta.valor = "") Then
        MENSAJE_ = "Ingrese un valor adecuado para la Fecha de Produccion"
        VERIFICO_ = False
        GoTo SALIR
    End If
    
    If (CDate(TxtFchHasta.valor) < CDate(TxtFchDesde.valor)) Then
        MENSAJE_ = "Ingrese un valor adecuado para la Fecha de Produccion"
        VERIFICO_ = False
    End If
SALIR:
    If Not VERIFICO_ Then MsgBox MENSAJE_, vbCritical + vbOKOnly, xTitulo
    verificarDatos = VERIFICO_
End Function

Private Sub hallarDatosRestantes(RESUMIDO As Boolean, DETALLADO As Boolean)
    Dim A As Integer
    Dim DETTAR_ As String
    
    If fg(0).Rows = fg(0).FixedRows Then Exit Sub
    
    ' Se halla el detalle de las Tareas
    hallarTareas
    
    lbl(1).Caption = "Tareas de Linea"
    CentrarFrm FraProgreso
    FraProgreso.Visible = True
    
    PgBar.Min = 0
    PgBar.Max = fg(0).Rows - 1
    Me.MousePointer = vbHourglass
    
    For A = 1 To fg(0).Rows - 1
        On Error GoTo SIGUIENTE
        If INTERRUMPIR_ Then GoTo SALIR
        
        FraProgreso.Refresh
        PgBar.Value = A
        
        RstTareas.Filter = adFilterNone
        RstTareas.Filter = "idctr = " & NulosN(fg(0).TextMatrix(A, 20)) _
                                    & " And corr = " & NulosN(fg(0).TextMatrix(A, 21))
        fg(0).TextMatrix(A, 15) = Format(RstTareas.RecordCount, "00")
        
        If DETALLADO Then
            RstTareas.MoveFirst
            DETTAR_ = NulosC(RstTareas("destar"))
            RstTareas.MoveNext
            While Not RstTareas.EOF
                DETTAR_ = DETTAR_ & " + " & NulosC(RstTareas("destar"))
                RstTareas.MoveNext
            Wend
            fg(0).TextMatrix(A, 16) = DETTAR_
        End If
SIGUIENTE:
    Next A
SALIR:
    Me.MousePointer = vbDefault
    FraProgreso.Visible = False
    GRID_AGRUPAR fg(0), 21
    INTERRUMPIR_ = False
End Sub

Private Sub hallarTareas()
    Dim xRs As New ADODB.Recordset
    Dim DETTAR_ As String
    
    cSQL = "SELECT pro_controltardettar.idctr, pro_controltardettar.corr, pro_tareas.descripcion AS destar " _
            + vbCr + "FROM pro_controltardettar LEFT JOIN pro_tareas ON pro_controltardettar.idtar = pro_tareas.id " _
            + vbCr + "Where (((pro_controltardettar.activo) = -1)) " _
            + vbCr + "ORDER BY pro_controltardettar.corr, pro_controltardettar.orden;"
    
    RST_Busq RstTareas, cSQL, xCon
End Sub

Private Sub generarConsulta(RESUMIDO As Boolean, DETALLADO As Boolean)
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    Dim cLINEA As String
    Dim cSUPERVISOR As String
    Dim cPERSONAL As String
    
    Me.MousePointer = vbHourglass
    
    If RESUMIDO Then
        cLINEA = GENERAR_SQL_ID(fg(3), 1, " AND pro_receta.iditem", "IN", True)
        cSUPERVISOR = GENERAR_SQL_ID(fg(4), 1, " AND pro_controltar.idres", "IN", True)
        
        cSQL = "SELECT alm_inventario.descripcion AS producto, pro_receta.codrec, pro_controltardet.numlote, pla_empleados.nombre AS nomres, pro_controltar.fchtra, mae_unidades.abrev, pro_controltardet.cant, pro_controltardet.horini AS horinilinea, pro_controltardet.horfin AS horfinlinea, pro_receta.iditem, pro_controltardet.idrec, pro_controltar.idres, pro_controltardet.idctr, pro_controltardet.corr, CNUMPER.numper " _
                + vbCr + "FROM ((pro_controltar LEFT JOIN pro_emp ON pro_controltar.idres = pro_emp.id) LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id) LEFT JOIN ((((pro_tareas RIGHT JOIN (mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN (pro_receta RIGHT JOIN pro_controltardet ON pro_receta.id = pro_controltardet.idrec) ON alm_inventario.id = pro_receta.iditem) ON mae_unidades.id = pro_controltardet.idunimed) ON pro_tareas.id = pro_controltardet.idtar) LEFT JOIN pro_controltardetgr ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) LEFT JOIN pla_empleados AS pla_empleados_1 ON pro_controltardetgr.idper = pla_empleados_1.id) LEFT JOIN " _
                + vbCr + "( " _
                + vbCr + "SELECT pro_controltardetgr.idctr, pro_controltardetgr.corr, Sum(1) AS numper " _
                + vbCr + "FROM pro_controltardetgr LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id " _
                + vbCr + "GROUP BY pro_controltardetgr.idctr, pro_controltardetgr.corr, pro_controltardetgr.activo " _
                + vbCr + "Having (((pro_controltardetgr.activo) = -1)) " _
                + vbCr + ") " _
                + vbCr + "AS CNUMPER ON (pro_controltardet.corr = CNUMPER.corr) AND (pro_controltardet.idctr = CNUMPER.idctr)) ON pro_controltar.id = pro_controltardet.idctr " _
                + vbCr + "GROUP BY alm_inventario.descripcion, pro_receta.codrec, pro_controltardet.numlote, pla_empleados.nombre, pro_controltar.fchtra, mae_unidades.abrev, pro_controltardet.cant, pro_controltardet.horini, pro_controltardet.horfin, pro_receta.iditem, pro_controltardet.idrec, pro_controltar.idres, pro_controltardet.idctr, pro_controltardet.corr, CNUMPER.numper, pro_controltardet.tipo, pro_controltardetgr.activo " _
                + vbCr + "HAVING (((pro_controltar.fchtra)>=CDate('" & TxtFchDesde.valor & "') And (pro_controltar.fchtra)<=CDate('" & TxtFchHasta.valor & "')) AND ((pro_controltardet.tipo)=3) AND ((pro_controltardetgr.activo)=-1)) " & cLINEA & cSUPERVISOR _
                + vbCr + "ORDER BY pro_controltardet.idctr, pro_controltardet.corr;"
            
    End If
    
    If DETALLADO Then
        cLINEA = GENERAR_SQL_ID(fg(3), 1, " AND pro_receta.iditem", "IN", True)
        cSUPERVISOR = GENERAR_SQL_ID(fg(4), 1, " AND pro_controltar.idres", "IN", True)
        cPERSONAL = GENERAR_SQL_ID(fg(5), 1, " AND pro_controltardetgr.idper", "IN", True)
        
        cSQL = "SELECT alm_inventario.descripcion AS producto, pro_receta.codrec, pro_controltardet.numlote, pla_empleados.nombre AS nomres, pro_controltar.fchtra, mae_unidades.abrev, pro_controltardet.cant, pro_controltardet.horini AS horinilinea, pro_controltardet.horfin AS horfinlinea, pla_empleados_1.nombre AS nomper, pro_controltardetgr.canpro, pro_controltardetgr.horini AS horiniper, pro_controltardetgr.horfin AS horfinper, pro_receta.iditem, pro_controltardet.idrec, pro_controltar.idres, pro_controltardetgr.idper, pro_controltardet.idctr, pro_controltardet.corr, CNUMPER.numper " _
                + vbCr + "FROM ((pro_controltar LEFT JOIN pro_emp ON pro_controltar.idres = pro_emp.id) LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id) LEFT JOIN ((((pro_tareas RIGHT JOIN (mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN (pro_receta RIGHT JOIN pro_controltardet ON pro_receta.id = pro_controltardet.idrec) ON alm_inventario.id = pro_receta.iditem) ON mae_unidades.id = pro_controltardet.idunimed) ON pro_tareas.id = pro_controltardet.idtar) LEFT JOIN pro_controltardetgr ON (pro_controltardet.idctr = pro_controltardetgr.idctr) AND (pro_controltardet.corr = pro_controltardetgr.corr)) LEFT JOIN pla_empleados AS pla_empleados_1 ON pro_controltardetgr.idper = pla_empleados_1.id) LEFT JOIN " _
                + vbCr + "( " _
                + vbCr + "SELECT pro_controltardetgr.idctr, pro_controltardetgr.corr, Sum(1) AS numper " _
                + vbCr + "FROM pro_controltardetgr LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id " _
                + vbCr + "GROUP BY pro_controltardetgr.idctr, pro_controltardetgr.corr, pro_controltardetgr.activo " _
                + vbCr + "HAVING (((pro_controltardetgr.activo)=-1)) " _
                + vbCr + ") " _
                + vbCr + "AS CNUMPER ON (pro_controltardet.idctr = CNUMPER.idctr) AND (pro_controltardet.corr = CNUMPER.corr)) ON pro_controltar.id = pro_controltardet.idctr " _
                + vbCr + "WHERE (((pro_controltar.fchtra)>=CDate('" & TxtFchDesde.valor & "') And (pro_controltar.fchtra)<=CDate('" & TxtFchHasta.valor & "')) AND ((pro_controltardet.tipo)=3) AND ((pro_controltardetgr.activo)=-1)) " & cLINEA & cSUPERVISOR & cPERSONAL _
                + vbCr + "ORDER BY pro_controltardet.idctr, pro_controltardet.corr;"
    End If
    
    RST_Busq xRs, cSQL, xCon
    
    ' se llenan los datos de la consulta
    DEFINIR_RST_TMP RstResumido, xRs
    CARGAR_RST_TMP RstResumido, xRs
    
    Me.MousePointer = vbDefault
    Set xRs = Nothing
End Sub

Private Sub llenarDatos(RESUMIDO As Boolean, DETALLADO As Boolean)
    Dim A As Integer
    
    fg(0).Rows = 1
    DoEvents
    If RstResumido.State = 0 Then Exit Sub
    If RstResumido.RecordCount = 0 Then Exit Sub
    
    RstResumido.MoveFirst
    lbl(1).Caption = "Registro de Lineas"
    CentrarFrm FraProgreso
    FraProgreso.Visible = True
    PgBar.Min = 0
    PgBar.Max = RstResumido.RecordCount
    Me.MousePointer = vbHourglass
    For A = 1 To RstResumido.RecordCount
        FraProgreso.Refresh
'        DoEvents
        PgBar.Value = A
        
        If INTERRUMPIR_ Then GoTo SALIR
        
        fg(0).Rows = fg(0).Rows + 1
        fg(0).TextMatrix(A, 1) = NulosC(RstResumido("producto"))
        fg(0).TextMatrix(A, 2) = NulosC(RstResumido("codrec"))
        fg(0).TextMatrix(A, 3) = NulosC(RstResumido("numlote"))
        fg(0).TextMatrix(A, 4) = NulosC(RstResumido("nomres"))
        fg(0).TextMatrix(A, 5) = Format(RstResumido("fchtra"), "dd/mm/yyyy")
        fg(0).TextMatrix(A, 6) = NulosC(RstResumido("abrev"))
        fg(0).TextMatrix(A, 7) = Format(NulosN(RstResumido("cant")), "0.00")
        fg(0).TextMatrix(A, 8) = Format(RstResumido("horinilinea"), "HH:mm")
        fg(0).TextMatrix(A, 9) = Format(RstResumido("horfinlinea"), "HH:mm")
        fg(0).TextMatrix(A, 10) = Format(RstResumido("numper"), "00")
        fg(0).TextMatrix(A, 17) = NulosN(RstResumido("iditem"))
        fg(0).TextMatrix(A, 18) = NulosN(RstResumido("idres"))
        fg(0).TextMatrix(A, 20) = NulosN(RstResumido("idctr"))
        fg(0).TextMatrix(A, 21) = NulosN(RstResumido("corr"))
        
        If DETALLADO Then
            fg(0).TextMatrix(A, 11) = NulosC(RstResumido("nomper"))
            fg(0).TextMatrix(A, 12) = Format(NulosN(RstResumido("canpro")), "0.00")
            fg(0).TextMatrix(A, 13) = Format(RstResumido("horiniper"), "HH:mm")
            fg(0).TextMatrix(A, 14) = Format(RstResumido("horfinper"), "HH:mm")
            fg(0).TextMatrix(A, 19) = NulosN(RstResumido("idper"))
        End If
        
        RstResumido.MoveNext
    Next A
    
SALIR:
    Me.MousePointer = vbDefault
    FraProgreso.Visible = False
    INTERRUMPIR_ = False
End Sub

Private Sub configurarGrid(RESUMIDO As Boolean, DETALLADO As Boolean)
    If RESUMIDO Then
        fg(0).ColWidth(11) = 0
        fg(0).ColWidth(12) = 0
        fg(0).ColWidth(13) = 0
        fg(0).ColWidth(14) = 0
        fg(0).ColWidth(16) = 0
    End If
    If DETALLADO Then
        fg(0).ColWidth(11) = 1800
        fg(0).ColWidth(12) = 825
        fg(0).ColWidth(13) = 765
        fg(0).ColWidth(14) = 780
        fg(0).ColWidth(16) = 2505
    End If
End Sub

Sub EXPORTAR()
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    Dim TITULO_ As String
    
    TITULO_ = "REGISTRO DE LINEAS"

    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, fg(0), TITULO_, "", "", TITULO_
    Set X_EXPORT = Nothing
    MousePointer = vbDefault
    Exit Sub
error:
    MousePointer = vbDefault
    SHOW_ERROR Name, "Exportar"
End Sub

Private Sub iniciarCampos()
    Dim MES_ As Integer
    Dim ANIO_ As Integer
    
    CARGO = False
    
    Set fg(3).DataSource = Nothing
    Set fg(4).DataSource = Nothing
    Set fg(5).DataSource = Nothing
    'Se inicializa:
    fg(0).Rows = 1
    'datos para clientes
    GRID_COMBOLIST fg(3), 2
    fg(3).Editable = flexEDKbdMouse
    'datos para productos
    GRID_COMBOLIST fg(4), 2
    fg(4).Editable = flexEDKbdMouse
    'datos para Ordenes de Compra
    GRID_COMBOLIST fg(5), 2
    fg(5).Editable = flexEDKbdMouse
    'datos para fechas
    TxtFchDesde.valor = Date
    TxtFchHasta.valor = Date
    ' datos para el reporte Simple
    fg(0).AllowUserResizing = flexResizeColumns
    fg(0).AutoSearch = flexSearchFromTop
    fg(0).ExplorerBar = flexExSortShow
    fg(0).SelectionMode = flexSelectionByRow
    fg(0).ForeColorSel = &H80000005
    fg(0).BackColorSel = &H80&
    ' Se ocultan las columnas correspondientes
    fg(0).ColWidth(17) = 0
    fg(0).ColWidth(18) = 0
    fg(0).ColWidth(19) = 0
    fg(0).ColWidth(20) = 0
    fg(0).ColWidth(21) = 0
    fg(0).FrozenCols = 5
    
    fg(3).ColWidth(1) = 0
    fg(4).ColWidth(1) = 0
    fg(5).ColWidth(1) = 0
    
    Opt(0).Value = True
    
    AGREGANDO_ = False
    INTERRUMPIR_ = False
End Sub

Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim xCampos() As String
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String
    
    If Index = 3 Then ' Lineas
        ReDim xCampos(1, 4) As String
        Dim nTitulo As String
        Dim xRsAux As New ADODB.Recordset
        
        Set xRs = Nothing
       
        cSQL = "SELECT alm_inventario.descripcion, pro_receta.iditem, alm_inventario.activo " _
            + vbCr + "FROM (pro_controltardet INNER JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) INNER JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id " _
            + vbCr + "GROUP BY alm_inventario.descripcion, pro_receta.iditem, alm_inventario.activo;"
        
        RST_Busq xRs, cSQL, xCon
        
        'descripcion                        'campo                           'tamaño                    'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "6000":    xCampos(0, 3) = "C"
                
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), "Buscando " & nTitulo, "descripcion", "descripcion", Principio

        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        If AGREGANDO_ Then
            fg(Index).Rows = fg(Index).Rows + 1
            fg(Index).Select fg(Index).Rows - 1, 1
        End If

        fg(Index).TextMatrix(fg(Index).Row, 1) = NulosN(xRs("iditem"))
        fg(Index).TextMatrix(fg(Index).Row, 2) = NulosC(xRs("descripcion"))
    End If
    
    If Index = 4 Then ' Supervisores
        ReDim xCampos(2, 4) As String
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Apellidos y Nombres":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":                xCampos(1, 1) = "id":        xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
                
        ' Se verifica que no se agregue una receta ya existente
        nSQLId = GENERAR_SQL_ID(fg(4), 1, " AND pro_emp.id", "NOT IN", True)
        
        cSQL = "SELECT pro_emp.id, pro_emp.sup, pro_emp.prog, pro_emp.res, pla_empleados.nombre " _
            + vbCr + "FROM (pro_emp LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id " _
            + vbCr + "Where (((pro_empdet.idfun) = 3)) " & nSQLId _
            + vbCr + "GROUP BY pro_emp.id, pro_emp.sup, pro_emp.prog, pro_emp.res, pla_empleados.nombre " _
            + vbCr + "Having (((pla_empleados.nombre) Is Not Null)) " _
            + vbCr + "ORDER BY pla_empleados.nombre;"
            
        nTitulo = "Buscando Personal Encargado"
                
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        If AGREGANDO_ Then
            fg(Index).Rows = fg(Index).Rows + 1
            fg(Index).Select fg(Index).Rows - 1, 1
        End If
        
        fg(Index).TextMatrix(fg(Index).Row, 1) = NulosN(xRs("id"))
        fg(Index).TextMatrix(fg(Index).Row, 2) = NulosC(xRs("nombre"))
    End If
    
    If Index = 5 Then ' Personal
        If Opt(0).Value Then
            MsgBox "Esta opcion no disponible para el Tipo de Consulta Resumido", vbCritical + vbOKOnly, xTitulo
            Exit Sub
        End If
        
        ReDim xCampos(4, 4) As String
        
        xCampos(0, 0) = "Num. Doc.":            xCampos(0, 1) = "numdoc":       xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
        xCampos(1, 0) = "Apellidos y Nombres":  xCampos(1, 1) = "nombre":       xCampos(1, 2) = "4000":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
        xCampos(2, 0) = "Fch. Ing.":            xCampos(2, 1) = "fching":       xCampos(2, 2) = "1000":     xCampos(2, 3) = "C":    xCampos(2, 4) = "C"
        xCampos(3, 0) = "Area":                 xCampos(3, 1) = "area":         xCampos(3, 2) = "2000":     xCampos(3, 3) = "C":    xCampos(3, 4) = "C"
                
        ' generar la lista de personal para no considerar en la lista
        nSQLId = GENERAR_SQL_ID(fg(5), 1, " AND pla_empleados.id", "NOT IN", True)
        
        ' generar la consulta
        cSQL = "SELECT pla_empleados.id AS idemp, pla_empleados.nombre, pla_empleados.numdoc, pla_empleados.fching, pla_empleados.fchcese, mae_area.descripcion AS area " _
            + vbCr + "FROM ((pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN mae_area ON pla_empleados.idarea = mae_area.id " _
            + vbCr + "WHERE (((pla_empleados.numdoc) Is Not Null And (pla_empleados.numdoc)<>'') AND ((pla_empleados.fchcese) Is Null Or (pla_empleados.fchcese)>=CDate('" & Date & "')) AND ((pro_empdet.idfun)=6)) " & nSQLId _
            + vbCr + "ORDER BY pla_empleados.nombre;"
            
        nTitulo = "Buscando Personal"
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
                      
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        If AGREGANDO_ Then
            fg(Index).Rows = fg(Index).Rows + 1
            fg(Index).Select fg(Index).Rows - 1, 1
        End If
        
        fg(Index).TextMatrix(fg(Index).Row, 1) = NulosN(xRs("idemp"))
        fg(Index).TextMatrix(fg(Index).Row, 2) = NulosC(xRs("nombre"))
    End If
    
    AGREGANDO_ = False
    Set xRs = Nothing
End Sub

Private Sub Fg_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case 3, 4, 5
            If KeyCode = vbKeyInsert Then ' Agregar
                menu00_Click
            End If
            
            If KeyCode = vbKeyDelete Then ' Eliminar
                menu01_Click
            End If
        Case Else
            Exit Sub
    End Select
End Sub

Private Sub Fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    INDICE_ = Index
    If Button <> 2 Then Exit Sub
    Select Case Index
        Case 3, 4, 5
            PopupMenu menu
        Case Else
            Exit Sub
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        INTERRUMPIR_ = True ' interrumpir
    End If
End Sub

Private Sub Form_Load()
    iniciarCampos
End Sub

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub
    
    If Me.Width <= 13200 Then Me.Width = 13200
    If Me.Height <= 2850 Then Me.Height = 2850
        
    ' Se dimensiona el contenido
    Frame6.Width = Me.Width - 90
    Frame6.Height = Me.Height - 795
    
    fg(0).Width = Frame6.Width - 105
    fg(0).Height = Frame6.Height - 975
End Sub

Private Sub menu00_Click()
    If fg(INDICE_).Rows > 2 Then fg(INDICE_).TopRow = fg(INDICE_).Rows - 2
    AGREGANDO_ = True
    Fg_CellButtonClick INDICE_, fg(INDICE_).Rows - 1, 1
End Sub

Private Sub menu01_Click()
    If fg(INDICE_).Row < fg(INDICE_).FixedRows Then Exit Sub
    fg(INDICE_).RemoveItem fg(INDICE_).Row
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        If verificarDatos Then
            Buscar
        End If
    End If
    
    If Button.Index = 5 Then
        EXPORTAR
    End If
    
    If Button.Index = 8 Then
        Unload Me
    End If
End Sub
