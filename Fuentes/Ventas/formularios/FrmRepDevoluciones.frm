VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRepDevoluciones 
   Caption         =   "Ventas  -  Reporte de Devoluciones"
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
               Picture         =   "FrmRepDevoluciones.frx":0000
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepDevoluciones.frx":0544
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepDevoluciones.frx":08D6
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepDevoluciones.frx":0A5A
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepDevoluciones.frx":0EAE
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepDevoluciones.frx":0FC6
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepDevoluciones.frx":150A
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepDevoluciones.frx":1A4E
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepDevoluciones.frx":1B62
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepDevoluciones.frx":1C76
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepDevoluciones.frx":20CA
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepDevoluciones.frx":2236
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepDevoluciones.frx":277E
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepDevoluciones.frx":2A98
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
      Begin VB.Frame Frame16 
         Caption         =   "[Fech. Doc.]"
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
         Cols            =   23
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmRepDevoluciones.frx":2E2A
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
         Index           =   1
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
         FormatString    =   $"FrmRepDevoluciones.frx":30BA
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
         Index           =   2
         Left            =   5280
         TabIndex        =   15
         ToolTipText     =   "Buscar Supervisor"
         Top             =   75
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
         FormatString    =   $"FrmRepDevoluciones.frx":3117
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
         Index           =   3
         Left            =   8430
         TabIndex        =   16
         ToolTipText     =   "Buscar Supervisor"
         Top             =   60
         Width           =   4815
         _cx             =   8493
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
         FormatString    =   $"FrmRepDevoluciones.frx":3176
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
Attribute VB_Name = "FrmRepDevoluciones"
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
    generarConsulta
End Sub

Private Function verificarDatos() As Boolean
    Dim VERIFICO_ As Boolean
    Dim MENSAJE_ As String
    
    VERIFICO_ = True
    If (Not IsDate(TxtFchDesde.Valor) Or Not IsDate(TxtFchHasta.Valor)) Then
        MENSAJE_ = "Ingrese un valor adecuado para la Fecha de Produccion"
        VERIFICO_ = False
        GoTo SALIR
    End If
    
    If (CDate(TxtFchHasta.Valor) < CDate(TxtFchDesde.Valor)) Then
        MENSAJE_ = "La fecha Hasta no puede ser mayor que la fecha Desde"
        VERIFICO_ = False
    End If
    
SALIR:
    If Not VERIFICO_ Then MsgBox MENSAJE_, vbCritical + vbOKOnly, xTitulo
    verificarDatos = VERIFICO_
End Function

Private Function GENERAR_SQL_ID_RST(Rst As ADODB.Recordset, nDesc As String, _
                            nCampo As String, Optional nTipoIn As String = "IN", _
                            Optional fEsNumero As Boolean = True) As String
    Dim nSQL As String
    Dim k&
    nSQL = ""
    
    If Rst.RecordCount = 0 Then Exit Function Else Rst.MoveFirst
    
    While Not Rst.EOF
        If Trim(CStr(Rst("" & nDesc & ""))) <> "" Then
            If fEsNumero = True Then
                nSQL = nSQL & NulosN(Rst("" & nDesc & "")) & ","
            Else
                nSQL = nSQL & "'" & NulosC(Rst("" & nDesc & "")) & "',"
            End If
        End If
        Rst.MoveNext
    Wend
    
    If nSQL <> "" Then nSQL = " " & nCampo & " " & nTipoIn & " (" + Left(nSQL, Len(nSQL) - 1) & ") "
        
    GENERAR_SQL_ID_RST = nSQL
End Function

Private Sub generarConsulta()
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    Dim cITEM As String
    Dim cCLIENTE As String
    Dim cMOTDEV As String
    
    Me.MousePointer = vbHourglass
         
    cITEM = GENERAR_SQL_ID(fg(1), 1, " AND vta_ventasdet.iditem", "IN", True)
    cCLIENTE = GENERAR_SQL_ID(fg(2), 1, " AND vta_ventas.idcli", "IN", True)
    cMOTDEV = GENERAR_SQL_ID(fg(3), 1, " AND vta_ventas.idmotdev", "IN", True)
    
    With fg(0)
        .Rows = 1
        
        ' Consulta de devoluciones
        cSQL = "SELECT vta_ventas.id, vta_ventas.idcli, mae_cliente.nombre AS descli, mae_cliente.numruc, vta_ventas.fchreg, vta_ventas.fchdoc, vta_ventas.tipdoc, mae_documento.abrev AS desdoc, vta_ventasdet.iditem, alm_inventario.descripcion, vta_ventas.numreg, [vta_ventas].[numser] & ' ' & [vta_ventas].[numdoc] AS numdoc, vta_ventas.idalm, alm_almacenes.descripcion AS desalm, vta_ventas.iddocref AS idfacref, [vta_ventas_1].[numser] & ' ' & [vta_ventas_1].[numdoc] AS numfacref, vta_ventas_1.fchdoc AS fchfacref, vta_ventasdet.canpro, mae_unidades.abrev, vta_ventasdet.preuni, vta_ventasdet.imptot, vta_ventas.idmon, mae_moneda.simbolo AS desmon, vta_ventas.idmotdev, mae_motivodevolucion.descripcion AS desmotdev, vta_ventas.desmotdev AS desmotdevotr " _
            + vbCr + "FROM (((((((((vta_ventas LEFT JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_documento AS mae_documento_1 ON vta_ventas.idtipdocref = mae_documento_1.id) LEFT JOIN alm_inventario ON vta_ventasdet.iditem = alm_inventario.id) LEFT JOIN alm_almacenes ON vta_ventas.idalm = alm_almacenes.id) LEFT JOIN vta_ventas AS vta_ventas_1 ON vta_ventas.iddocref = vta_ventas_1.id) LEFT JOIN mae_unidades ON vta_ventasdet.idunimed = mae_unidades.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN mae_motivodevolucion ON vta_ventas.idmotdev = mae_motivodevolucion.id " _
            + vbCr + "WHERE (((vta_ventas.idmotnotcre)=4) AND ((vta_ventas.fchdoc)>=CDate('" & TxtFchDesde.Valor & "') And (vta_ventas.fchdoc)<=CDate('" & TxtFchHasta.Valor & "'))) " & cITEM & cCLIENTE & cMOTDEV _

        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Me.MousePointer = vbDefault: Exit Sub
        
        xRs.MoveFirst
        For A = 1 To xRs.RecordCount
            .Rows = .Rows + 1
            .TextMatrix(A, 1) = NulosN(xRs("id"))
            .TextMatrix(A, 2) = NulosN(xRs("iditem"))
            .TextMatrix(A, 3) = NulosN(xRs("idcli"))
            .TextMatrix(A, 4) = NulosN(xRs("idmotdev"))
            .TextMatrix(A, 5) = NulosN(xRs("idmon"))
            .TextMatrix(A, 6) = NulosN(xRs("idfacref"))
            .TextMatrix(A, 7) = NulosN(xRs("tipdoc"))
            
            .TextMatrix(A, 8) = NulosC(xRs("descli"))
            .TextMatrix(A, 9) = NulosC(xRs("numruc"))
            .TextMatrix(A, 10) = NulosC(xRs("descripcion"))
            .TextMatrix(A, 11) = Format(NulosC(xRs("fchdoc")), FORMAT_DATE)
            .TextMatrix(A, 12) = NulosC(xRs("numreg"))
            .TextMatrix(A, 13) = NulosC(xRs("numdoc"))
            .TextMatrix(A, 14) = NulosC(xRs("numfacref"))
            .TextMatrix(A, 15) = Format(NulosC(xRs("fchfacref")), FORMAT_DATE)
            .TextMatrix(A, 16) = Format(NulosN(xRs("canpro")), FORMAT_CANTIDAD)
            .TextMatrix(A, 17) = NulosC(xRs("abrev"))
            .TextMatrix(A, 18) = Format(NulosN(xRs("preuni")), FORMAT_CANTIDAD)
            .TextMatrix(A, 19) = Format(NulosN(xRs("imptot")), FORMAT_CANTIDAD)
            .TextMatrix(A, 20) = NulosC(xRs("desmon"))
            .TextMatrix(A, 21) = NulosC(xRs("desmotdev"))
            .TextMatrix(A, 22) = NulosC(xRs("desmotdevotr"))
                        
            xRs.MoveNext
        Next A
        configurarGrid

    End With

    Me.MousePointer = vbDefault
    Set xRs = Nothing
End Sub

Private Sub configurarGrid()
    fg(0).ColWidth(1) = 0
    fg(0).ColWidth(2) = 0
    fg(0).ColWidth(3) = 0
    fg(0).ColWidth(4) = 0
    fg(0).ColWidth(5) = 0
    fg(0).ColWidth(6) = 0
    fg(0).ColWidth(7) = 0
    
    fg(0).FrozenCols = 13
    fg(0).RowHeight(0) = 300
End Sub

Sub EXPORTAR()
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    Dim TITULO_ As String
    
    TITULO_ = "REPORTE DE DEOLUCIONES"

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
    
    Set fg(1).DataSource = Nothing
    Set fg(2).DataSource = Nothing
    Set fg(3).DataSource = Nothing
    'Se inicializa:
    fg(0).Rows = 1
    'datos para clientes
    GRID_COMBOLIST fg(1), 2
    fg(1).Editable = flexEDKbdMouse
    'datos para productos
    GRID_COMBOLIST fg(2), 2
    fg(2).Editable = flexEDKbdMouse
    'datos para Ordenes de Compra
    GRID_COMBOLIST fg(3), 2
    fg(3).Editable = flexEDKbdMouse
    'datos para fechas
    TxtFchDesde.Valor = Date
    TxtFchHasta.Valor = Date
    ' datos para el reporte Simple
    fg(0).AllowUserResizing = flexResizeColumns
    fg(0).AutoSearch = flexSearchFromTop
    fg(0).ExplorerBar = flexExSortShow
    fg(0).SelectionMode = flexSelectionByRow
    fg(0).ForeColorSel = &H80000005
    fg(0).BackColorSel = &H80&
    
    fg(1).ColWidth(1) = 0
    fg(2).ColWidth(1) = 0
    fg(3).ColWidth(1) = 0
        
    AGREGANDO_ = False
    INTERRUMPIR_ = False
    
    configurarGrid
End Sub

Private Sub fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim xCampos() As String
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String
    
    If Index = 1 Then ' Lineas
        ReDim xCampos(1, 4) As String
        Dim nTitulo As String
        Dim xRsAux As New ADODB.Recordset
        
        Set xRs = Nothing
        
        nSQLId = GENERAR_SQL_ID(fg(Index), 1, " AND alm_inventario.id", "NOT IN", True)

        cSQL = "SELECT alm_inventario.descripcion, pro_receta.iditem " _
            + vbCr + "FROM pro_receta LEFT JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id " _
            + vbCr + "WHERE (((alm_inventario.activo)=-1) AND ((pro_receta.prirec)=1)) " & nSQLId
        
        RST_Busq xRs, cSQL, xCon
        
        'descripcion                        'campo                           'tamaño                    'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "6000":    xCampos(0, 3) = "C"
                
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), "Buscando " & nTitulo, "descripcion", "descripcion", Principio

        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub

        fg(Index).TextMatrix(fg(Index).Row, 1) = NulosN(xRs("iditem"))
        fg(Index).TextMatrix(fg(Index).Row, 2) = NulosC(xRs("descripcion"))
    End If
    
    If Index = 2 Then ' Cliente
        ReDim xCampos(2, 3) As String
        Set xRs = Nothing
        
        xCampos(0, 0) = "Nombre":          xCampos(0, 1) = "nombre":     xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
        xCampos(1, 0) = "Nº R.UC.":        xCampos(1, 1) = "numruc":     xCampos(1, 2) = "1300":   xCampos(1, 3) = "C"
        
        ' Se verifica que no se agregue una receta ya existente
        nSQLId = GENERAR_SQL_ID(fg(2), 1, " AND mae_cliente.id", "NOT IN", True)
        
        cSQL = "SELECT mae_cliente.id, mae_cliente.nombre, mae_cliente.numruc " _
               + vbCr + "FROM mae_cliente " _
               + vbCr + "WHERE mae_cliente.nombre <> '' " & nSQLId _
                
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), "Buscando Clientes", "nombre", "nombre", Principio
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        fg(Index).TextMatrix(fg(Index).Row, 1) = xRs("id")
        fg(Index).TextMatrix(fg(Index).Row, 2) = xRs("nombre")
    End If
    
    If Index = 3 Then ' Mot. Devoluciones
        ReDim xCampos(2, 3) As String
        Set xRs = Nothing
        
        xCampos(0, 0) = "Id":       xCampos(0, 1) = "id":            xCampos(0, 2) = "1000":   xCampos(0, 3) = "N"
        xCampos(1, 0) = "Motivo":   xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "5000":   xCampos(1, 3) = "C"
        
        ' Se verifica que no se agregue una receta ya existente
        nSQLId = GENERAR_SQL_ID(fg(3), 1, " AND mae_motivodevolucion.id", "NOT IN", True)
        
        cSQL = "SELECT mae_motivodevolucion.id, mae_motivodevolucion.descripcion " _
               + vbCr + "FROM mae_motivodevolucion " _
               + vbCr + "WHERE mae_motivodevolucion.id is not null " & nSQLId _
                
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), "Buscando MOtivos de Devolucion", "descripcion", "descripcion", Principio
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        fg(Index).TextMatrix(fg(Index).Row, 1) = xRs("id")
        fg(Index).TextMatrix(fg(Index).Row, 2) = xRs("descripcion")
    End If
    
        
    If fg(Index).Row = fg(Index).Rows - 1 Then
        fg(Index).Rows = fg(Index).Rows + 1
        fg(Index).Select fg(Index).Rows - 1, 2
        fg(Index).TopRow = fg(Index).Rows - 1
    End If
        
    AGREGANDO_ = False
    Set xRs = Nothing
End Sub

Private Sub fg_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case 1, 2, 3
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

Private Sub fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    INDICE_ = Index
    If Button <> 2 Then Exit Sub
    Select Case Index
        Case 1, 2, 3
            PopupMenu Menu
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

Private Sub menu00_Click() ' Agregar
    If fg(INDICE_).Rows > 2 Then fg(INDICE_).TopRow = fg(INDICE_).Rows - 2
    AGREGANDO_ = True
    fg_CellButtonClick INDICE_, fg(INDICE_).Rows - 1, 1
End Sub

Private Sub menu01_Click() ' Eliminar
    If fg(INDICE_).Row < fg(INDICE_).FixedRows Then Exit Sub
    fg(INDICE_).RemoveItem fg(INDICE_).Row
    
    If fg(INDICE_).Rows = fg(INDICE_).FixedRows Then fg(INDICE_).Rows = fg(INDICE_).Rows + 1
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
