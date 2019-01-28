VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmConsControlTarea 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Producción - Consulta de Tareas"
   ClientHeight    =   7665
   ClientLeft      =   105
   ClientTop       =   600
   ClientWidth     =   11805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   11805
   Begin VB.Frame FraProgreso 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   2835
      TabIndex        =   22
      Top             =   3720
      Visible         =   0   'False
      Width           =   5760
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   90
         TabIndex        =   23
         Top             =   345
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tareas"
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
         Index           =   1
         Left            =   1185
         TabIndex        =   26
         Top             =   75
         Width           =   585
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   90
         TabIndex        =   25
         Top             =   75
         Width           =   1035
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Index           =   2
         Left            =   4140
         TabIndex        =   24
         Top             =   75
         Width           =   1530
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5745
         X2              =   5745
         Y1              =   -90
         Y2              =   4800
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -150
         X2              =   5895
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   2
         X1              =   -60
         X2              =   6360
         Y1              =   675
         Y2              =   690
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   3
         X1              =   0
         X2              =   15
         Y1              =   15
         Y2              =   5070
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Align           =   2  'Align Bottom
      Height          =   5160
      Left            =   0
      TabIndex        =   8
      Top             =   2505
      Width           =   11805
      _cx             =   20823
      _cy             =   9102
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmConsControlTarea.frx":0000
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   609
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
         Left            =   4860
         Top             =   90
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
               Picture         =   "FrmConsControlTarea.frx":0211
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsControlTarea.frx":0755
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsControlTarea.frx":0AE7
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsControlTarea.frx":0C6B
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsControlTarea.frx":10BF
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsControlTarea.frx":11D7
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsControlTarea.frx":171B
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsControlTarea.frx":1C5F
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsControlTarea.frx":1D73
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsControlTarea.frx":1E87
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsControlTarea.frx":22DB
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsControlTarea.frx":2447
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsControlTarea.frx":298F
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsControlTarea.frx":2CA9
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2235
      Left            =   0
      TabIndex        =   2
      Top             =   255
      Width           =   11805
      Begin VB.CheckBox chk_obs 
         Caption         =   "Mostrar sólo Observados"
         Height          =   195
         Left            =   6810
         TabIndex        =   27
         Top             =   660
         Width           =   2325
      End
      Begin VB.CommandButton cb 
         Height          =   240
         Index           =   0
         Left            =   6345
         Picture         =   "FrmConsControlTarea.frx":303B
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   225
         Width           =   225
      End
      Begin VB.CheckBox chk_grupo 
         Caption         =   "Buscar Grupos"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5310
         TabIndex        =   16
         Top             =   660
         Width           =   1425
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Consulta"
         Height          =   750
         Left            =   2100
         TabIndex        =   5
         Top             =   135
         Width           =   1515
         Begin VB.OptionButton opt_consulta 
            Caption         =   "&Detallado"
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   3
            Top             =   450
            Width           =   975
         End
         Begin VB.OptionButton opt_consulta 
            Caption         =   "&Resumen"
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   0
            Top             =   225
            Value           =   -1  'True
            Width           =   1185
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   1245
         Index           =   0
         Left            =   45
         TabIndex        =   6
         ToolTipText     =   "Buscar Personal / Nº Grupo"
         Top             =   930
         Width           =   3855
         _cx             =   6800
         _cy             =   2196
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
         FormatString    =   $"FrmConsControlTarea.frx":316D
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
         Height          =   1245
         Index           =   1
         Left            =   7905
         TabIndex        =   7
         ToolTipText     =   "Buscar Productos"
         Top             =   930
         Width           =   3855
         _cx             =   6800
         _cy             =   2196
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
         AllowUserResizing=   0
         SelectionMode   =   1
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
         FormatString    =   $"FrmConsControlTarea.frx":31D5
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
      Begin VB.Frame Frame5 
         Caption         =   "Agrupar Por"
         Height          =   750
         Left            =   3720
         TabIndex        =   9
         Top             =   135
         Width           =   1515
         Begin VB.OptionButton opt_grupo 
            Caption         =   "x &Producto"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   1
            Top             =   225
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opt_grupo 
            Caption         =   "x Tarea"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   4
            Top             =   450
            Width           =   855
         End
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFec1 
         Height          =   300
         Left            =   645
         TabIndex        =   11
         Top             =   225
         Width           =   1365
         _ExtentX        =   2408
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
         Valor           =   "25/09/2007"
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFec2 
         Height          =   300
         Left            =   645
         TabIndex        =   12
         Top             =   585
         Width           =   1365
         _ExtentX        =   2408
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
         Valor           =   "25/09/2007"
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   1245
         Index           =   2
         Left            =   3975
         TabIndex        =   15
         ToolTipText     =   "Buscar Tareas"
         Top             =   930
         Width           =   3855
         _cx             =   6800
         _cy             =   2196
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
         AllowUserResizing=   0
         SelectionMode   =   1
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
         FormatString    =   $"FrmConsControlTarea.frx":324A
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
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   0
         Left            =   5745
         MaxLength       =   12
         TabIndex        =   18
         Text            =   "txt_cb(0)"
         Top             =   195
         Width           =   870
      End
      Begin VB.Label lbl_cb 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cb(0)"
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
         Index           =   0
         Left            =   6600
         TabIndex        =   21
         Top             =   195
         Width           =   5085
      End
      Begin VB.Label lbl_cod 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cod(0)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   0
         Left            =   7950
         TabIndex        =   20
         Top             =   195
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label lbl_cb_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Area"
         Height          =   195
         Index           =   0
         Left            =   5340
         TabIndex        =   19
         Top             =   300
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         Height          =   195
         Left            =   90
         TabIndex        =   14
         Top             =   330
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   675
         Width           =   465
      End
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu Menu2_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu2_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu2_3 
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "FrmConsControlTarea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-- ALMACENAR LOS TOTALES DE TODA LA CONSULTA
'--ARR_TMP(?,4)= Arr_Totales_cols() As Double '--ALMACENAR TOTALES POR TODAS LAS FILAS
'--ARR_TMP(?,3)= Arr_Totales_col() As Double     '--ALMACENAR TOTALES POR COLUMNA, SE LIMPIA DESPUES DE CAMBIO DE GRUPO


Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE
'--DE LA IMPRESION
Dim T_RPT_PERIODO As String '--PERIODO DEL REPORTE
Dim T_RPT_TITULO As String  '--TITULO DE REPORTE
'------------
Dim ARR_ANYO() As String    '--ARRAY DE AÑOS SELECCIONADOS
Dim ARR_XX() As String      '--SE CARGARA CUANDO SE CARGA EL FORMULARIO Y CUANDO SE CAMBIE EL ESTILO(MES, TRIMESTRE,SEMESTRE)

Dim ARR_TMP(3, 1) As String '--0::PROGRAMADO=>> 0::TOTAL,1::TOTAL GEN
                            '--1::TEORICO=>> 0::TOTAL,1::TOTAL GEN
                            '--2::REAL=>> 0::TOTAL,1::TOTAL GEN
                            '--3::DIF=>> 0::TOTAL,1::TOTAL GEN

Dim Q_TOTAL_ANYO As Integer '--INDICA LA CANTIDAD DE AÑOS DE BUSQUEDA,
                            '--EJ. 2004,2005 => Q_TOTAL_ANYO = 2
                            '--EJ. 2004,2005,2006 => Q_TOTAL_ANYO = 3
                            
Dim Q_COL_FILA As Integer   '--INDICA LA CANTIDAD DE COLUMNAS ANTES DE LOS MESES
                            '--EJ. IDCLI,CLIENTE => Q_COL_FILA=2
                            '--    IDCLI,ID_PTO_VTA,CLIENTE,PTO_VENTA => Q_COL_FILA=4
                            
                            
Dim Q_COL_FILA_ULTIMO As Integer '--INDICA LA CANTIDAD DE COLUMNAS ADICIONALES QUE SE COLOCARAN DESPUES DEL TOTAL
                            
Dim Q_POS_MES_INICIO As Integer '--INDICA LA POSICION INICIAL DE LA COLUMNA DEL PRIMER MES, NO CAMBIA
                            '--EJ. Q_POS_MES_INICIO = Q_COL_FILA +1

Dim Q_POS_MES As Integer    '--INDICA LA POSICION DEL MES, ESTO CAMBIA
                            '--UTIL PARA COLOCAR LOS DATOS EN EL GRID

Dim Q_COL_FILA_OCULTA As Integer '--INDICA LAS COLUMNAS QUE CONTENDRAN LOS ID'S, ESTOS SE OCULTARAN
                                '-- -1 NO SE OCULTA, <> -1 SE PROCEDE A ACULTAR
                                'EJ. CLIENTE  vta_ventas.idcli,
                                    'PUNTO DE VENTA vta_guia.idpunven
                                    'PRODUCTO   alm_inventario.tippro
                                    'ITEM       alm_inventario.id
                                    'EMPLEADO   vta_ventas.idven

Dim Q_POSICION_TOTAL  As Integer '--INDICA LA POCISION DE LA COLUMNA DONDE SE COLOCARA EL NOMBRE DEL TOTAL Y TOTAL_GRL
                                 '--OBTENDRA VALOR EN fGenerarConsulta()

Dim Q_COL_COMPARAR_GRUPO As Integer '--INDICA LA COLUMNA PARA COMPARAR EL GRUPO
                                    '--OBTENDRA VALOR EN fGenerarConsulta()

'------------------------------
Dim Q_COL_GRUPO_ADD As Integer  '--ADICIONAR DATOS AL GRID EN EL GRUPO (EJ. Q_COL_GRUPO_ADD=2 =>> NOMBRE_GRUPO|COLUM1|COLUM2)
                                '--FNUCIONA SI Q_COL_GRUPO_ADD<>-1

Dim N_CAMPO_GRUPO_ADD As String '--INDICA EL NOMBRE DEL CAMPO A COMPARAR PARA AGREGAR AL LA FILA DEL GRUPO DEPENDE DE Q_COL_GRUPO_ADD
'------------------------------
                                
                                
Dim Q_COL_GRUPO_INICIO      As Integer  '--INDICA EL INICIO DE LA COLUMNA DEL GRUPO,
Dim Q_COL_GRUPO_TERMINA     As Integer  '--INDICA EL TERMINO DE LA COLUMNA DEL GRUPO, UNE LAS CELDAS DE [Q_COL_GRUPO_INICIO] HASTA [Q_COL_GRUPO_TERMINA]

'------------

Dim Q_COL_ARR_TOTAL As Integer  '--NOS INDICA EL TOTAL DE COLUMNAS QUE VA A CONTENER LOS ACUMULADOS
                                '--OBTENDRA VALOR EN fValidarConsulta()
                                '--SI SEL MES: ENE, FEB, MAR => Q_COL_ARR_TOTAL= 2
                                '--SI SEL TRI: ENE-MAR, ABR-JUN => Q_COL_ARR_TOTAL= 1 OBS: SE INICIA DESDE POS=0

Dim F_ES_COMPRA As Boolean '--INDICA SI ES COMPRA O VENTA
                            '--TRUE::ES COMPRA, FALSE::ES VENTA

Dim ID_PROGRAMA As String
Dim ID_RECETA As String
Dim TIPO_VENTANA As e_PROGRAMA
Dim ESTILO_VISTA As Integer
'-------
Dim nSQLValor_FONDO           As String '--AMACENA EL VALOR PARA COMPARAR
Dim nSQLValor_FONDO_COLOR     As Long '--AMACENA EL VALOR DEL COLOR PARA EL FONDO DE LA FILA
Dim F_CAMIAR_FONDO          As Boolean  '--FALSE::SE CONSERVA EL FONDO ACTUAL, TRUE::CAMBIA DE FONDO
Dim Q_COL_COMPARAR_FONDO    As Integer  '--INDICA LA COLUMNA DEL RECORDSET QUE DEBERA DE COMPARAR PARA CAMBIAR DE FONDO
                                        '-- -1=NO HACER NADA
'------------

Private Sub pConsultar()
'    On Error GoTo error
    Dim rst_select As New ADODB.Recordset
    Dim nSQLSelect As String '--RECIBIR LA CONSULTA
    '--CARGANDO RUTAS DE LOS AÑOS SELECCIONADOS
    
    If fValidarConsulta() = False Then Exit Sub
    nSQLValor_FONDO = ""
    BAND_INTERRUMPIR = False
    
    '--CONFIGURAR LA PRESENTACION DE LA CONSULTA
    LimpiarGrid Me.Fg1, False, 1
    '--ENTRAR SOLO UNA VEZ
    nSQLSelect = fGenerarConsulta()
    pConfigurarGrilla
        
    '--LIMPIAR ARRAY
    Limpiar_ARRAY_TOTAL True
    '----
    Me.MousePointer = vbHourglass
    DoEvents
    
    '------------------------------------------------
    If nSQLSelect = "" Then GoTo SALIR
    PosicionarProgBar
    DoEvents
    '--CARGADO EL RST
    RST_Busq rst_select, nSQLSelect, xCon
   '--------------------------------------
    pCargarDatosGrid rst_select
   '--------------------------------------
   '
SALIR:
    FraProgreso.Visible = False
    Set rst_select = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    FraProgreso.Visible = False
    Set rst_select = Nothing
    SHOW_ERROR Me.Name, "pConsultar"
    
End Sub

Private Sub pCargarDatosGridFondo(RST_ORIGEN As ADODB.Recordset, _
                                        X_ROW1 As Long, X_COL1 As Integer, _
                                        X_ROW2 As Long, X_COL2 As Integer)

    ''--PONER COLOR FONDO
    If Q_COL_COMPARAR_FONDO = -1 Then Exit Sub
        If NulosN(Fg1.TextMatrix(X_ROW1, 1)) = e_ESTADO_ROW_GRID.Fila_grupo Then
            '--SI SE DESEA PONER COLOR AL GRUPO
            'GRID_COLOR_FONDO Fg1, X_ROW1, X_COL1, X_ROW2, X_COL2, RGB(0, 185, 185)
        ElseIf NulosN(Fg1.TextMatrix(X_ROW1, 1)) = e_ESTADO_ROW_GRID.Fila_Total Then
        ElseIf NulosN(Fg1.TextMatrix(X_ROW1, 1)) = e_ESTADO_ROW_GRID.Fila_Total_grl Then
        ElseIf NulosN(Fg1.TextMatrix(X_ROW1, 1)) = e_ESTADO_ROW_GRID.Fila_en_Blanco Then
        Else
           If nSQLValor_FONDO = "" Then
                '--se coloca la opcion "-" para considerar los nulos
                nSQLValor_FONDO = NulosC(RST_ORIGEN.Fields(Q_COL_COMPARAR_FONDO)) & "-"
                nSQLValor_FONDO_COLOR = &HFDFFFF
                F_CAMIAR_FONDO = False
            End If
    
            If nSQLValor_FONDO = NulosC(RST_ORIGEN.Fields(Q_COL_COMPARAR_FONDO)) & "-" Then
                nSQLValor_FONDO_COLOR = nSQLValor_FONDO_COLOR
            Else
                nSQLValor_FONDO = NulosC(RST_ORIGEN.Fields(Q_COL_COMPARAR_FONDO)) & "-"
                If F_CAMIAR_FONDO = True Then
                    nSQLValor_FONDO_COLOR = &HFDFFFF
                    F_CAMIAR_FONDO = False
                Else
                    nSQLValor_FONDO_COLOR = &HE0FEFE
                    F_CAMIAR_FONDO = True
                End If
            End If
            GRID_COLOR_FONDO Fg1, X_ROW1, X_COL1, X_ROW2, X_COL2, nSQLValor_FONDO_COLOR
        End If


End Sub

Private Function pCargarDatosGrid(RST_ORIGEN As ADODB.Recordset)
                                         
    '--FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
    Dim BAND_ADD_REG As Boolean
    
    
    BAND_ADD_REG = True
    
    If RST_ORIGEN.RecordCount > 0 Then
        RST_ORIGEN.MoveFirst
    Else
        Exit Function
    End If
    PgBar.Min = 0
    PgBar.Max = RST_ORIGEN.RecordCount
    
    While Not RST_ORIGEN.EOF
    
    DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then Exit Function
        '---------------------------------------------------------
        pCompararGrupo RST_ORIGEN, BAND_ADD_REG, Q_COL_COMPARAR_GRUPO
        
        If RST_ORIGEN.Bookmark <> 1 Then ADD_REG Fg1

        '--CARGAR A LA GRILLA
        pCargarDatosGridArrayTmp RST_ORIGEN, Fg1.Rows - 1
        
        '---------------------------------------------------------
        '---------------------------------------------------------
        ''--PONER COLOR FONDO
        If Q_COL_COMPARAR_FONDO <> -1 Then pCargarDatosGridFondo RST_ORIGEN, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
            
        '---------------------------------------------------------
        '---------------------------------------------------------
        RST_ORIGEN.MoveNext
'        --PONER TOTALES AL FINAL DE LA GRILLA
        
        If RST_ORIGEN.EOF Then
            pCargarDatosGridAddTotales BAND_ADD_REG, "Total:"
            Select Case ESTILO_VISTA
            Case 0, 1, 2, 4, 5, 8, 9
            Case Else
                pCargarDatosGridAddTotales True, "Tot Gen:", True
            End Select
        Else
            PgBar.Value = CLng(RST_ORIGEN.Bookmark)
        End If
    Wend
    
    '------

End Function



Private Sub pCompararGrupo(RST_ORIGEN As ADODB.Recordset, _
                            BAND_ADD_REG As Boolean, _
                            Optional Q_COL_COMPARAR As Integer = -1)
                            
    '--FUNCION QUE NOS PERMITE ARMAR LOS GRUPOS
    '--COMPARA CUANDO CAMBIAR DE GRUPO
    Dim RST_TEPM_1 As New ADODB.Recordset
    Dim N_GRUPO_ADD As String
    Dim Q_POS As Integer
    
    '---------------------------------------------------------
    If Q_COL_COMPARAR = -1 Then
        If RST_ORIGEN.Bookmark = 1 Then ADD_REG Fg1, Fila_Ninguno
        Exit Sub
    End If
    '---------------------------------------------------------
    If Q_COL_COMPARAR = -1 Then Q_COL_COMPARAR = Q_COL_COMPARAR_GRUPO
    
    If Q_COL_GRUPO_ADD <> -1 Then
        If NulosC(N_CAMPO_GRUPO_ADD) <> "" Then
            For Q_POS = 1 To Q_COL_GRUPO_ADD
                If LCase(RST_ORIGEN.Fields(Q_COL_COMPARAR + Q_POS).Name) = UCase(N_CAMPO_GRUPO_ADD) Then
                    N_GRUPO_ADD = Format(NulosN(RST_ORIGEN.Fields(Q_COL_COMPARAR + Q_POS)), FORMAT_MONTO) + " " + N_GRUPO_ADD
                Else
                    N_GRUPO_ADD = RST_ORIGEN.Fields(Q_COL_COMPARAR + Q_POS) & "  " + N_GRUPO_ADD
                End If
            Next Q_POS
        End If
        N_GRUPO_ADD = "  =>>   " + N_GRUPO_ADD
    End If
    
    If RST_ORIGEN.Bookmark = 1 Then
        '--SE CARGA EN fGenerarConsulta() Q_COL_COMPARAR_GRUPO
        ADD_REG Fg1, Fila_grupo
        UNIR_CELDAS Fg1, Fg1.Rows - 1, Q_COL_GRUPO_INICIO, Fg1.Rows - 1, Q_COL_GRUPO_TERMINA, INICIO_GRUPO & NulosC(RST_ORIGEN.Fields(Q_COL_COMPARAR)) & N_GRUPO_ADD, flexAlignLeftCenter:
        FORMATO_CELDA Fg1, Fg1.Rows - 1, Q_COL_GRUPO_INICIO
        '--------------
        ADD_REG Fg1, Fila_Ninguno
        UNIR_CELDAS Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, " ", flexAlignLeftCenter
        
        nSQLValor_FONDO = ""
    Else
    
        Set RST_TEPM_1 = RST_ORIGEN.Clone
        RST_TEPM_1.Bookmark = RST_ORIGEN.Bookmark
        RST_TEPM_1.MovePrevious
        
        If RST_TEPM_1.Fields(Q_COL_COMPARAR) <> RST_ORIGEN.Fields(Q_COL_COMPARAR) Then
            '--cargar datos de total
            pCargarDatosGridAddTotales BAND_ADD_REG, "Total:"
            
            '--poner la fila en blanco, agrupado
            ADD_REG Fg1, Fila_en_Blanco
            UNIR_CELDAS Fg1, Fg1.Rows - 1, IIf(Q_COL_FILA_OCULTA = -1, 1, Q_COL_FILA_OCULTA + 1), Fg1.Rows - 1, Fg1.Cols - 1, " ", flexAlignLeftCenter
            
            Limpiar_ARRAY_TOTAL
            
            ADD_REG Fg1, Fila_grupo
            UNIR_CELDAS Fg1, Fg1.Rows - 1, Q_COL_GRUPO_INICIO, Fg1.Rows - 1, Q_COL_GRUPO_TERMINA, INICIO_GRUPO & RST_ORIGEN.Fields(Q_COL_COMPARAR) & N_GRUPO_ADD, flexAlignLeftCenter
            FORMATO_CELDA Fg1, Fg1.Rows - 1, Q_COL_GRUPO_INICIO
            
            '--inicializando el color del fondo
            nSQLValor_FONDO = ""

        End If
    End If

SALIR:
    Set RST_TEPM_1 = Nothing


End Sub

Private Function pCargarDatosGridArrayTmp(RST_ORIGEN As ADODB.Recordset, _
                                         Q_ROW As Long)
                                         
    '--FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
    
    Dim Q_INCREMENTO_X_COL As Integer   '--SERVIRA PARA POSICIONAR LA PRIMERA COLUMNA DE ENERO DE UN AÑO
    Dim Q_POS_MES_TOTAL  As Integer     '--POSICIONA A LA COLUMNA DEL TOTAL X AÑO
    Dim Q_POS As Integer
    Dim Q_CAMPO As Integer
    Dim vStrCampo As String
    
    '--IDENTIFICAR LA POSICION DE INICIO DE MES(ENERO)
    Q_INCREMENTO_X_COL = 0
    Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
    '-----------
    
    DoEvents
    
    For Q_CAMPO = 0 To RST_ORIGEN.Fields.Count - 1
        If BAND_INTERRUMPIR = True Then Exit Function
        vStrCampo = RST_ORIGEN.Fields(Q_CAMPO).Name
        
        Select Case LCase(vStrCampo)
            Case "acumulado", "total"
                Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO)
            Case "horini", "horfin"
                Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(RST_ORIGEN.Fields(vStrCampo), FORMAT_HORA_SIN_SEGUNDO)
            Case "fchtra"
                Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(RST_ORIGEN.Fields(vStrCampo), FORMAT_DATE)
            
            Case Else
                '--AGREGAR LOS DEMAS DATOS
                Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = RST_ORIGEN.Fields(vStrCampo) & ""
        End Select
    Next
End Function


Private Sub pImprimir()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    X_PRINT.Imprimir_x_VSFlexGrid Fg1, T_RPT_TITULO + " ", "", T_RPT_PERIODO, True, True
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"

End Sub


Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)

    If Row = 0 Then Exit Sub
   
   If NulosC(fg(Index).TextMatrix(Row, Col)) = "" Then fg(Index).TextMatrix(Row, 1) = ""
    
End Sub

Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If validar_letras(KeyAscii) = False Then
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If

End Sub

Private Sub Fg1_DblClick()
    Fg1_KeyDown 13, 0
End Sub

Private Sub Fg1_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> 13 Then Exit Sub
'    If Fg1.Rows = 1 Then Exit Sub
'    If Fg1.Row = 0 Or Fg1.Row = Fg1.Rows - 1 Or Fg1.Col = 0 Or Fg1.Col = 1 Or Fg1.Col = 2 Or Fg1.Col = Fg1.Cols - 1 Then
'        MsgBox "Selecione una Celda Correcta..", vbInformation, "Mensaje"
'        Exit Sub
'    End If
'    If txt(5).Text = "" Or IsNumeric(txt(5).Text) = False Then
'        MsgBox "Ingrese un número a mostrar", vbInformation, "Mensaje..."
'        txt(5).SetFocus
'        Exit Sub
'    End If
'    If IsNumeric(Fg1.TextMatrix(Fg1.Row, Fg1.Col)) = False Then
'        MsgBox "La celda no es numérico", vbInformation, "Mensaje..."
'        Exit Sub
'    End If
    
'    With FrmAnalizaPrecio_Item
'        .RECIBE_ID_ITEM Fg1.TextMatrix(Fg1.Row, 1), Fg1.TextMatrix(1, Fg1.Col), ARR_TMP(), F_ES_COMPRA
'        .Show 1
'    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo error
    CentrarFrm Me
    LimpiaText txt_cb
    LimpiaText lbl_cb
    '--FORMATO DE LAS GRILLAS
    GRID_COMBOLIST fg(0), 2:        fg(0).Tag = fg(0).FormatString
    GRID_COMBOLIST fg(1), 2:        fg(1).Tag = fg(1).FormatString
    GRID_COMBOLIST fg(2), 2:        fg(2).Tag = fg(2).FormatString
    '-------
    TxtFec1.Valor = CDate("01/01/" + CStr(Year(Date)))
    TxtFec2.Valor = Date
    fGenerarConsulta
    pConfigurarGrilla
    
    Exit Sub
error:
    SHOW_ERROR
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    BAND_INTERRUMPIR = True
    Erase ARR_TMP
End Sub



'------
Private Function fValidarConsulta() As Boolean
    '--FUNCION QUE VALIDARA LA CONSULTA
    '--DE LA FECHA ES NULL
    If TxtFec1.Valor = "" Or TxtFec2.Valor = "" Then
        MsgBox "Ingrese una fecha", vbExclamation, xTitulo
        If TxtFec1.Valor = "" Then TxtFec1.SetFocus Else TxtFec2.SetFocus
        Exit Function
    End If
    If CDate(TxtFec1.Valor) > CDate(TxtFec2.Valor) Then
        MsgBox "La fecha inicial es superior al Final", vbExclamation, xTitulo
        TxtFec1.SetFocus
        Exit Function
    End If
    fValidarConsulta = True
End Function

Private Function fGenerarConsulta() As String
    '--FUNCION QUE NOS PERMITIRA GENERAR LA CONSULTA DE ACUERDO A LO QUE SELECCIONE EL USUARIO
    '--
    Dim nSQLSelect As String            '--CONSULTA GENERAL, ESTO PERMITIRA HACER LA CONSULTA
    
    Dim nSQLFecha As String     '--almacenar el intervalo de fechas
    Dim nSQLProducto As String  '--almacenar los id's de productos
    Dim nSQLPersonal As String  '--almacenara los id's del personal
    Dim nSQLTarea As String     '--almacenara los id's de las tareas
    Dim nSQLArea As String      '--almacena el id del area
    Dim nSQLObs As String       '--almacena si se muestran solo los observados
    Dim mTipoConsulta As Integer '--DEL TIPO DE CONSULTA, SE FORMARA EL ENCABEZADO DEL GRID
    
    '--de la fecha
    If CDate(TxtFec1.Valor) < CDate(TxtFec2.Valor) Then
        nSQLFecha = " ( pro_controltar.fchtra BETWEEN CDATE ('" + TxtFec1.Valor + "') AND CDATE('" + TxtFec2.Valor + "') ) "
        T_RPT_PERIODO = " Del: " + CStr(TxtFec1.Valor) + " Al: " + CStr(TxtFec2.Valor)
    Else
        nSQLFecha = " pro_controltar.fchtra = CDATE('" + TxtFec1.Valor + "') "
         T_RPT_PERIODO = "Al: " + CStr(TxtFec2.Valor)
   End If
    '--de los productos
    nSQLProducto = GENERAR_SQL_ID(fg(1), 1, " AND alm_inventario.id", "IN")
    '--de las tareas
    nSQLTarea = GENERAR_SQL_ID(fg(2), 1, " AND pro_controltardet.idtar ", "IN")
    '--del area
    If NulosN(lbl_cod(0).Caption) <> 0 Then nSQLArea = " AND pro_controltar.idarea = " & NulosN(lbl_cod(0).Caption)
    '--solo los observados
    If chk_obs.Value = 1 Then nSQLObs = " AND pro_controltardet.observado = -1 "
    '--GENERAR LA CONSULTA SEGUN CONDICIONES
    Dim nSQLValor As String
    Dim nSQLCampos As String
    Dim nSQLWhere As String
    Dim nSQLFrom As String
    Dim nSQLGroupBy As String
    Dim nSQLOrderBy As String
    
    mTipoConsulta = fEstiloConsulta()
    Q_COL_COMPARAR_FONDO = -1
    
    Select Case mTipoConsulta
        Case 0, 1 '--resumido - producto; resumido - tarea
        
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 5:        Q_POSICION_TOTAL = 3:
            '-------------------
            '--ADICIONAR DATOS AL GRID EN EL GRUPO (NOMBRE_GRUPO|COLUM1|COLUM2)
            Q_COL_GRUPO_ADD = -1:   N_CAMPO_GRUPO_ADD = ""
            '-------------------
            Q_COL_GRUPO_INICIO = 1: Q_COL_GRUPO_TERMINA = 5
            
            If mTipoConsulta = 0 Then
                Q_COL_COMPARAR_GRUPO = 1
                Q_COL_COMPARAR_FONDO = 2 '--num lote
                T_RPT_TITULO = "RESUMEN DE TAREAS AGRUPADO POR PRODUCTO"
            Else
                Q_COL_COMPARAR_GRUPO = 3
                Q_COL_COMPARAR_FONDO = 2 '--num lote
                T_RPT_TITULO = "RESUMEN AGRUPADO POR TAREA"
            End If
            
            nSQLSelect = "SELECT vw.id , vw.producto,vw.numlote, vw.tarea,sum(vw.total) as acumulado ,vw.abrev " _
                + vbCr + " FROM ( "
            nSQLSelect = nSQLSelect _
                + vbCr + " SELECT alm_inventario.id ,IIf([alm_inventario].[descripcion] Is Not Null,[alm_inventario].[descripcion],'* ' & [pro_controltardet].[observacion]) AS producto, pro_controltardet.numlote, pro_tareas.descripcion AS tarea, Sum(pro_controltardet.cant) AS total, mae_unidades.abrev " _
                + vbCr + " FROM (pro_controltar INNER JOIN (alm_inventario RIGHT JOIN ((pro_controltardet INNER JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id " _
                + vbCr + " WHERE pro_controltardet.tipo=1 and " & nSQLFecha & nSQLArea & nSQLTarea & nSQLProducto _
                + vbCr + " GROUP BY alm_inventario.id ,IIf([alm_inventario].[descripcion] Is Not Null,[alm_inventario].[descripcion],'* ' & [pro_controltardet].[observacion]), pro_controltardet.numlote, pro_tareas.descripcion, pro_tareas.descripcion, mae_unidades.abrev " _
                + vbCr + " HAVING (((pro_controltardet.numlote)<>'' And (pro_controltardet.numlote) Is Not Null))"
            nSQLSelect = nSQLSelect _
                + vbCr + " UNION "
            nSQLSelect = nSQLSelect _
                + vbCr + " SELECT  alm_inventario.id ,IIf([alm_inventario].[descripcion] Is Not Null,[alm_inventario].[descripcion],'* ' & [pro_controltardet].[observacion]) AS producto, pro_controltardet.numlote, pro_tareas.descripcion AS tarea, Sum(pro_controltardetgr.cant) AS total, mae_unidades.abrev " _
                + vbCr + " FROM (pro_controltar INNER JOIN (alm_inventario RIGHT JOIN (((pro_controltardet INNER JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) INNER JOIN pro_controltardetgr ON (pro_controltardet.idctr = pro_controltardetgr.idctr) AND (pro_controltardet.corr = pro_controltardetgr.corr)) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id " _
                + vbCr + " WHERE pro_controltardet.tipo=2 and " & nSQLFecha & nSQLArea & nSQLTarea & nSQLProducto _
                + vbCr + " GROUP BY alm_inventario.id ,IIf([alm_inventario].[descripcion] Is Not Null,[alm_inventario].[descripcion],'* ' & [pro_controltardet].[observacion]), pro_controltardet.numlote, pro_tareas.descripcion, mae_unidades.abrev " _
                + vbCr + " HAVING (((pro_controltardet.numlote)<>'' And (pro_controltardet.numlote) Is Not Null))"
            nSQLSelect = nSQLSelect _
                + vbCr + " ) AS vw " _
                + vbCr + " GROUP BY vw.id, producto, vw.numlote,vw.tarea,vw.abrev " _
                + vbCr + " ORDER BY vw.producto, vw.numlote asc,vw.tarea"
                
                
           
        Case 2 '--detalle
        
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 12:        Q_POSICION_TOTAL = 9:        Q_COL_COMPARAR_GRUPO = -1
            '-------------------
            '--ADICIONAR DATOS AL GRID EN EL GRUPO (NOMBRE_GRUPO|COLUM1|COLUM2)
            Q_COL_GRUPO_ADD = -1:   N_CAMPO_GRUPO_ADD = ""
            '-------------------
            Q_COL_GRUPO_INICIO = -1: Q_COL_GRUPO_TERMINA = -1
            Q_COL_COMPARAR_FONDO = 4
            T_RPT_TITULO = "DETALLE DEL REGISTRO DE TAREAS "
            If chk_obs.Value = 1 Then T_RPT_TITULO = "DETALLE DEL REGISTRO DE TAREAS " & " - OBSERVADOS"
            
            nSQLCampos = "  [pro_controltardet].[idctr] & [pro_controltardet].[corr] AS codigo,pro_controltardet.numlote, pro_controltar.fchtra, mae_area.descripcion, IIf([pro_controltardet].[tipo]=1,[pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom],'Grupo Nº ' & [pro_controltardet].[idref]) AS nombres, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardet.horini, pro_controltardet.horfin, pro_controltardet.cant AS total, mae_unidades.abrev, IIf([pro_controltardet].[observado]=-1,'Si',' ') AS Obs, pro_controltardet.observacion "
            nSQLFrom = " ((pro_controltar INNER JOIN (alm_inventario RIGHT JOIN (((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr) LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id "
            nSQLWhere = nSQLFecha & nSQLArea & nSQLTarea & nSQLProducto & nSQLObs
            nSQLOrderBy = " pro_controltar.fchtra, IIf([pro_controltardet].[tipo]=1,[pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom],'Grupo Nº ' & [pro_controltardet].[idref]), pro_controltardet.horini;"
        
        
    End Select

    '--DE LA CANTIDAD DE COL
    Q_COL_FILA = Q_COL_FILA + 1
    Q_POS_MES_INICIO = Q_COL_FILA '--Q_COL_FILA + CAMPO_TOTAL
    
    '------------------------------------------

    '--GENERANDO LA CONSULTA
    If nSQLSelect = "" Then
        nSQLSelect = "SELECT " + nSQLCampos + _
        vbCr + " FROM " + nSQLFrom + _
        vbCr + " WHERE " + nSQLWhere + _
        vbCr + IIf(nSQLGroupBy <> "", " GROUP BY ", "") + nSQLGroupBy + _
        vbCr + " ORDER BY " + nSQLOrderBy
    End If

    '------------------------------------------------------------------------------------
    fGenerarConsulta = nSQLSelect
    
End Function

Private Sub Limpiar_ARRAY_TOTAL(Optional F_LIMPIA_TOT_GRL As Boolean = False)
    Dim k As Integer
    For k = 0 To UBound(ARR_TMP())
'        ARR_TMP(k, 3) = 0
'        If F_LIMPIA_TOT_GRL = True Then ARR_TMP(k, 4) = 0
    Next
                            
End Sub
'''
Private Sub pCargarDatosGridAddTotales(BAND_ADD_TOTAL As Boolean, _
                                            Nombre_total As String, _
                                            Optional Band_Total_gral As Boolean = False)
                
    Dim Q_MES As Integer
    '--AGREGA LOS TOTALES POR CADA GRUPO Y EL TOTAL GENERAL
    '--ACUMULA LOS TOTALES EN EL TOTAL GENERAL
    Dim X_ROW As Integer
''''''    'On Error Resume Next
''''''    X_ROW = Fg1.Rows
''''''    If BAND_ADD_TOTAL = True Then
''''''        '--AGREAGNDO NUEVA FILA
''''''        ADD_REG Fg1, IIf(Band_Total_gral = False, Fila_Total, Fila_Total_grl)
''''''
''''''        'PONIENDO LOS NOMBRES DE LOS TOTALES  Q_POSICION_TOTAL SE OBTIENE DE fGenerarConsulta()
''''''        Fg1.TextMatrix(X_ROW, Q_POSICION_TOTAL) = Nombre_total
''''''        FORMATO_CELDA Fg1, X_ROW, Q_POSICION_TOTAL
''''''    End If
''''''
''''''
''''''    '--ACUMULANDO LOS TOTALES GRLES
''''''    If Band_Total_gral = False Then
''''''        For Q_MES = 0 To Q_COL_ARR_TOTAL
''''''            ARR_TMP(Q_MES, 4) = NulosN(ARR_TMP(Q_MES, 4)) + NulosN(ARR_TMP(Q_MES, 3))
''''''        Next Q_MES
''''''        If Q_COL_FILA_ULTIMO <> -1 Then
''''''            ARR_TMP_1(0, 1) = NulosN(ARR_TMP_1(0, 1)) + NulosN(ARR_TMP_1(0, 0)) '--STOCK
''''''            ARR_TMP_1(1, 1) = NulosN(ARR_TMP_1(1, 1)) + NulosN(ARR_TMP_1(1, 0)) '--SALDO
''''''        End If
''''''    End If
''''''    '
'''''''--------------------------
''''''    Dim Q_INCREMENTO_X_COL As Integer   '--SERVIRA PARA POSICIONAR LA PRIMERA COLUMNA DE ENERO DE UN AÑO
''''''    Dim Q_POS_MES_TOTAL  As Integer     '--POSICIONA A LA COLUMNA DEL TOTAL X AÑO
''''''
''''''
''''''    '--IDENTIFICAR LA POSICION DE INICIO DE MES(ENERO)
''''''    Q_INCREMENTO_X_COL = 0
''''''    Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
''''''    '-----------
''''''
''''''    Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
''''''
''''''    For Q_MES = 0 To Q_COL_ARR_TOTAL
''''''        '--INTERRUMPIR EL PROCESO
''''''        If BAND_INTERRUMPIR = True Then Exit Sub
''''''        Fg1.TextMatrix(X_ROW, Q_POS_MES) = PONER_FORMATO(IIf(Band_Total_gral = False, ARR_TMP(Q_MES, 3), ARR_TMP(Q_MES, 4)), Band_Total_gral, Q_MES)
''''''        FORMATO_CELDA Fg1, X_ROW, Q_POS_MES
''''''        Q_POS_MES = Q_POS_MES + 1
''''''    Next Q_MES
''''''
''''''
''''''    If Q_COL_FILA_ULTIMO <> -1 Then
''''''        '--STOCK
''''''        Fg1.TextMatrix(X_ROW, Fg1.Cols - 2) = PONER_FORMATO(IIf(Band_Total_gral = False, ARR_TMP_1(0, 0), ARR_TMP_1(0, 1)), Band_Total_gral, Fg1.Cols - 2)
''''''        FORMATO_CELDA Fg1, X_ROW, Fg1.Cols - 2, RGB(128, 0, 0)
''''''        '--SALDO
''''''        Fg1.TextMatrix(X_ROW, Fg1.Cols - 1) = PONER_FORMATO(IIf(Band_Total_gral = False, ARR_TMP_1(1, 0), ARR_TMP_1(1, 1)), Band_Total_gral, Fg1.Cols - 1)
''''''        FORMATO_CELDA Fg1, X_ROW, Fg1.Cols - 1, vbRed
''''''    End If
''''''    Err.Clear
End Sub

Private Sub pConfigurarGrilla()
    '--PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA
    '--DE ACUERDO A LO QUE SE SELECCIONA
    Dim M_ANCHO_COL As Integer '--DEPENDERA DEL TIPO DE CONSULTA
                                   
    Dim k, j As Integer
    Dim mTipoConsulta As Integer
    
    Fg1.FrozenCols = 0
    
    M_ANCHO_COL = 0

    With Fg1
        '-----
        Fg1.Cols = Q_COL_FILA_OCULTA + Q_COL_FILA
                 
        Q_POS_MES = Q_POS_MES_INICIO
        
        '.FrozenCols = Q_POS_MES_INICIO - 1
        .ColWidth(0) = 200
        '--DATOS DE FILA
        
    mTipoConsulta = fEstiloConsulta()
    Select Case mTipoConsulta
        Case 0 '--resumido / X PRODUCTO
                .TextMatrix(0, 2) = "Producto":     .ColWidth(2) = 0:  .ColAlignment(2) = flexAlignLeftBottom:      .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftBottom
                .TextMatrix(0, 3) = "Nº Lote":      .ColWidth(3) = 1500:  .ColAlignment(3) = flexAlignLeftBottom:   .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftBottom
                .TextMatrix(0, 4) = "Tarea":        .ColWidth(4) = 4500:  .ColAlignment(4) = flexAlignLeftBottom:   .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftBottom
                .TextMatrix(0, 5) = "Total":        .ColWidth(5) = 1300:  .ColAlignment(5) = flexAlignRightBottom:  .Row = 0: .Col = 5: .CellAlignment = flexAlignRightCenter
                .TextMatrix(0, 6) = "U.M.":         .ColWidth(6) = 450:   .ColAlignment(6) = flexAlignCenterCenter: .Row = 0: .Col = 6: .CellAlignment = flexAlignCenterCenter
                
        Case 1 '--resumido / X num lote
                .TextMatrix(0, 2) = "Producto":     .ColWidth(2) = 4500:  .ColAlignment(2) = flexAlignLeftBottom:   .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftBottom
                .TextMatrix(0, 3) = "Nº Lote":      .ColWidth(3) = 1500:  .ColAlignment(3) = flexAlignLeftBottom:   .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftBottom
                .TextMatrix(0, 4) = "Tarea":        .ColWidth(4) = 0:  .ColAlignment(4) = flexAlignLeftBottom:      .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftBottom
                .TextMatrix(0, 5) = "Total":        .ColWidth(5) = 1300:  .ColAlignment(5) = flexAlignRightBottom:  .Row = 0: .Col = 5: .CellAlignment = flexAlignRightCenter
                .TextMatrix(0, 6) = "U.M.":         .ColWidth(6) = 450:   .ColAlignment(6) = flexAlignCenterCenter: .Row = 0: .Col = 6: .CellAlignment = flexAlignCenterCenter
                
        Case 2 '--detalle
                .TextMatrix(0, 2) = "Nº Lote":          .ColWidth(2) = 1200:      .ColAlignment(2) = flexAlignLeftBottom:         .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftBottom
                .TextMatrix(0, 3) = "Fecha":            .ColWidth(3) = 800:       .ColAlignment(3) = flexAlignCenterBottom:       .Row = 0: .Col = 3: .CellAlignment = flexAlignCenterBottom
                .TextMatrix(0, 4) = "Area":             .ColWidth(4) = 700:       .ColAlignment(4) = flexAlignLeftBottom:         .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftBottom
                .TextMatrix(0, 5) = "Personal / Grupo": .ColWidth(5) = 2000:      .ColAlignment(5) = flexAlignLeftBottom:         .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftBottom
                '--------
                .TextMatrix(0, 6) = "Tarea":            .ColWidth(6) = 2800:      .ColAlignment(6) = flexAlignLeftBottom:         .Row = 0: .Col = 6: .CellAlignment = flexAlignLeftBottom
                .TextMatrix(0, 7) = "Producto":         .ColWidth(7) = 3000:      .ColAlignment(7) = flexAlignLeftBottom:         .Row = 0: .Col = 7: .CellAlignment = flexAlignLeftBottom
                .TextMatrix(0, 8) = "H.Inicio":         .ColWidth(8) = 800:       .ColAlignment(8) = flexAlignCenterCenter:       .Row = 0: .Col = 8: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(0, 9) = "H.Final":          .ColWidth(9) = 800:       .ColAlignment(9) = flexAlignCenterCenter:       .Row = 0: .Col = 9: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(0, 10) = "Cant":            .ColWidth(10) = 850:      .ColAlignment(10) = flexAlignRightBottom:       .Row = 0: .Col = 10: .CellAlignment = flexAlignRightBottom
                .TextMatrix(0, 11) = "U.M.":            .ColWidth(11) = 500:      .ColAlignment(11) = flexAlignCenterCenter:      .Row = 0: .Col = 11: .CellAlignment = flexAlignCenterCenter
                .TextMatrix(0, 12) = "Obs":             .ColWidth(12) = 400:      .ColAlignment(12) = flexAlignLeftBottom:        .Row = 0: .Col = 12: .CellAlignment = flexAlignLeftBottom
                .TextMatrix(0, 13) = "Inf. Adicional":  .ColWidth(13) = 2000:     .ColAlignment(13) = flexAlignLeftBottom:        .Row = 0: .Col = 13: .CellAlignment = flexAlignLeftBottom
                '--ocultar la columna de obs si solo se muestran los observados
                If chk_obs.Value = 1 Then
                    .ColWidth(12) = 0
                    .ColWidth(13) = 4000
                End If
    End Select

        'If Q_COL_COMPARAR_GRUPO <> -1 Then .ColWidth(3) = 0
        
        '--DE LOS ID'S
        For k = 1 To Q_COL_FILA_OCULTA
            .TextMatrix(0, k) = "ID" + CStr(k):         .ColWidth(k) = 500
        Next
        If Q_COL_FILA_OCULTA <> -1 Then OCULTAR_COL Fg1, 1, Q_COL_FILA_OCULTA
        
'        If Q_COL_GRUPO_ADD <> -1 Then OCULTAR_COL Fg1, Q_COL_COMPARAR_GRUPO + 1, Q_COL_COMPARAR_GRUPO + Q_COL_GRUPO_ADD + 1
    
        
    End With
    DoEvents
End Sub


Private Function PONER_FORMATO(S_MONTO As Double, _
                        Optional Band_Total_gral As Boolean = False, _
                        Optional Q_POS As Integer = -1) As String
                        
    '--ESTA FUNCION CONVERTIRA AL FORMATO
    If S_MONTO = 0 Then
            PONER_FORMATO = "0.00"
        Exit Function
    End If
    
    PONER_FORMATO = Format(S_MONTO, FORMAT_MONTO)
    
    
    
    
End Function

Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    Dim xCampos() As String
    Dim nTitulo As String
    Dim nSQL As String
    Dim nSQLNotIn As String
    Dim Q_ROW As Long
    If Col <> 2 Then Exit Sub
    Select Case Index
        
        Case 1 '--producto
            ReDim xCampos(3, 4) As String
            xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "codpro":       xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
            xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "5000":    xCampos(1, 3) = "C"
            xCampos(2, 0) = "id":           xCampos(2, 1) = "id":           xCampos(2, 2) = "700":     xCampos(2, 3) = "N"
                          
            '--de los registros ya seleccionados
            nSQLNotIn = GENERAR_SQL_ID(fg(Index), 1, " AND alm_inventario.id", "NOT IN")
            
            If NulosC(fg(Index).TextMatrix(Row, Col)) <> "" Then
                nSQLNotIn = nSQLNotIn & " AND UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' "
            End If
            
            nSQL = "SELECT alm_inventario.id,alm_inventario.codpro, alm_inventario.descripcion " _
                + vbCr + " FROM alm_inventario " _
                + vbCr + " WHERE alm_inventario.id IN (SELECT pro_receta.iditem FROM pro_receta INNER JOIN pro_recetatar ON pro_receta.id = pro_recetatar.idrec;) " & nSQLNotIn _
                + vbCr + " ORDER BY alm_inventario.descripcion; "
            
            nTitulo = "Buscando Producto"
            
        Case 2 '--tarea
    
            ReDim xCampos(4, 4) As String
            xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "descripcion":     xCampos(0, 2) = "4500":    xCampos(0, 3) = "C"
            xCampos(1, 0) = "Abrev":        xCampos(1, 1) = "nomcorto":   xCampos(1, 2) = "2300":    xCampos(1, 3) = "C"
            xCampos(2, 0) = "Diverso":      xCampos(2, 1) = "diverso":    xCampos(2, 2) = "700":     xCampos(2, 3) = "C"
            xCampos(3, 0) = "Id":           xCampos(3, 1) = "id":         xCampos(3, 2) = "600":     xCampos(3, 3) = "N"
            
            '--de los registros ya seleccionados
            nSQLNotIn = GENERAR_SQL_ID(fg(Index), 1, " WHERE pro_tareas.id", "NOT IN")
            
            If NulosC(fg(Index).TextMatrix(Row, Col)) <> "" Then
                If nSQLNotIn = "" Then
                    nSQLNotIn = " WHERE "
                Else
                    nSQLNotIn = nSQLNotIn & " AND "
                End If
                nSQLNotIn = nSQLNotIn & " (UCASE(pro_tareas.descripcion) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' OR UCASE(pro_tareas.abrev) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' ) "
            End If
            
            nSQL = "SELECT pro_tareas.id, pro_tareas.codigo, pro_tareas.descripcion , pro_tareas.abrev AS nomcorto, mae_unidades.id AS idunimed, mae_unidades.abrev, IIf([pro_tareas].[diverso]=-1,'Si','No') AS diverso " _
                    + vbCr + " FROM mae_unidades RIGHT JOIN pro_tareas ON mae_unidades.id = pro_tareas.idunimed  " & nSQLNotIn
            
            nTitulo = "Buscando Tareas"
            
        Case Else
            Exit Sub
    End Select

    Dim xRs As New ADODB.Recordset
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    
    fg(Index).TextMatrix(fg(Index).Row, 1) = NulosC(xRs.Fields("id"))
    fg(Index).TextMatrix(fg(Index).Row, 2) = NulosC(xRs.Fields("descripcion"))
        
        
    If fg(Index).Row = fg(Index).Rows - 1 Then fg(Index).AddItem ""
    fg(Index).Row = fg(Index).Rows - 1: fg(Index).Col = 2
        
SALIR:

    Set xRs = Nothing

Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Fg_CellButtonClick(" + CStr(Index) + ")", True, "Error..."

End Sub

Private Sub Fg_DblClick(Index As Integer)
    Fg_CellButtonClick Index, fg(Index).Rows - 1, 2
End Sub

Private Sub Fg_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If fg(Index).Row = -2 Then Exit Sub
    Select Case KeyCode
        Case 45  'INSERTAR REGI
            fg(Index).AddItem ""
            fg(Index).Row = fg(Index).Rows - 1: fg(Index).Col = 2
        Case 46 'SUPRIMIR/DELETE
            If fg(Index).Rows - 1 >= 2 Then
                fg(Index).RemoveItem fg(Index).Row
                fg(Index).Row = fg(Index).Rows - 1: fg(Index).Col = 2
            Else
                LimpiarGrid fg(Index), True
                GRID_COMBOLIST fg(Index)
            End If
    End Select
End Sub


Private Function fEstiloConsulta() As Integer
    Dim mTipoEstilo As Integer
    If opt_consulta(0).Value = True Then '--RESUMEN
        If opt_grupo(0).Value = True Then mTipoEstilo = 0 '--X producto
        If opt_grupo(1).Value = True Then mTipoEstilo = 1 '--X lote
    Else '--DETALLE
        mTipoEstilo = 2 '--X detalle
    End If
    fEstiloConsulta = mTipoEstilo

End Function

Private Sub PosicionarProgBar()
    '--POSICIONAR EL PROGRESO DENTRO DEL FORMULARIO
    '    FraProgreso.Width = 5820
    FraProgreso.Left = (Me.Width - FraProgreso.Width) \ 2
    FraProgreso.Top = (Me.Height - FraProgreso.Height) \ 2
    Me.PgBar.Value = 1
    FraProgreso.Visible = True
End Sub


'----
Private Sub Fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Index = 0 Then PopupMenu Menu1
        If Index = 1 Then PopupMenu menu2
    End If
End Sub

'--DEL PRODUCTO
Private Sub Menu1_1_Click()
    Fg_DblClick 0
End Sub

Private Sub Menu1_3_Click()
    Fg_KeyDown 0, 46, 0
End Sub
'--DE LOS INSUMOS
Private Sub Menu2_1_Click()
    Fg_DblClick 1
End Sub

Private Sub Menu2_3_Click()
    Fg_KeyDown 1, 46, 0
End Sub
'--------
Private Sub pExportarExcel()
On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios

    Me.MousePointer = vbHourglass
    X_PRINT.VSFlexGrid_Exportar_MSExcel xCon, Fg1, T_RPT_TITULO + " ", T_RPT_PERIODO, , "Producción"
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportarExcel"
End Sub


Private Sub opt_consulta_Click(Index As Integer)
    If Index = 0 Then '--resumen
        habilitar opt_grupo, True
        opt_grupo(0).Value = True
        chk_obs.Enabled = False
    Else '--detalle
        opt_grupo(0).Value = False
        opt_grupo(1).Value = False
        habilitar opt_grupo, False
        chk_obs.Enabled = True
    End If
End Sub

'************************************************

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 5 Then pExportarExcel
    If Button.Index = 6 Then pImprimir
    If Button.Index = 8 Then
        BAND_INTERRUMPIR = True
        Unload Me
    End If
End Sub

'************************************************


'*******************************************************************************************

Private Sub cb_Click(Index As Integer)
    Dim xCampos() As String
    Dim nCampoBusca As String
    Dim nSQL As String
    Dim nTitulo As String
    On Error GoTo error
    Select Case Index
        Case 0 '--area
            nTitulo = "Buscando Area"
            nSQL = "SELECT mae_area.id, mae_area.descripcion AS nombre, mae_area.id AS cod, mae_area.id AS idarea " _
                + vbCr + " FROM pro_area INNER JOIN mae_area ON pro_area.idarea = mae_area.id; "
            
    End Select
    
    ReDim xCampos(2, 3) As String
    xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":           xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
    
    Dim RstTmp As New ADODB.Recordset
    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio

    If RstTmp.State = 0 Then GoTo SALIR
    If RstTmp.EOF = True Or RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index).Text = NulosC(RstTmp.Fields(0))  '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = NulosC(RstTmp.Fields(1)) '--NOMBRE
    lbl_cod(Index).Caption = NulosN(RstTmp.Fields(2)) '--CODIGO
    lbl_cb(Index).ToolTipText = NulosC(RstTmp.Fields(1))  '--NOMBRE
      
    Select Case Index
        Case 0
            fg(0).SetFocus
            
    End Select
SALIR:
    Set RstTmp = Nothing
Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub


Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cod(Index).Caption = ""
    End If
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If txt_cb(Index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Index <> 1 Then
            SendKeys vbTab
        Else
            If Fg1.Rows >= 2 Then
                Fg1.Row = 1: Fg1.Col = 1
            Else
                Fg1.Row = Fg1.Rows - 1: Fg1.Col = 1
            End If
            Fg1.SetFocus
        End If
        Exit Sub
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)

    If txt_cb(Index).Text = "" Then Exit Sub
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
        Case 0 '--area
            nSQL = "SELECT mae_area.id, mae_area.descripcion AS nombre, mae_area.id AS cod, mae_area.id AS idarea " _
                + vbCr + " FROM pro_area INNER JOIN mae_area ON pro_area.idarea = mae_area.id; "
        
        Case Else
            Exit Sub
            
    End Select

    If xCon.State = 0 Then GoTo SALIR
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo SALIR

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cb(Index).Text = NulosC(RstTmp.Fields(0))   '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = NulosC(RstTmp.Fields(1)) '--NOMBRE
        lbl_cod(Index).Caption = NulosN(RstTmp.Fields(2)) '--CODIGO
        lbl_cb(Index).ToolTipText = NulosC(RstTmp.Fields(1)) '--NOMBRE
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cod(Index).Caption = ""
    End If
    
    Set RstTmp = Nothing
    Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "txt_cb_Validate (" + CStr(Index) + ")"
    Exit Sub
SALIR:
    Set RstTmp = Nothing
    txt_cb(Index).Text = ""
End Sub

'****************************************************************************************


