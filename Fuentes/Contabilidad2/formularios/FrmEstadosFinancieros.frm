VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form FrmEstadosFinancieros 
   Caption         =   "Contabilidad - Estados Financieros"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11985
   LinkTopic       =   "Form2"
   ScaleHeight     =   7440
   ScaleWidth      =   11985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   6210
      Begin VB.TextBox txtnumdigitos 
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "txtnumdigitos"
         Top             =   480
         Width           =   735
      End
      Begin ComCtl2.UpDown UpdDigitos 
         Height          =   285
         Left            =   1695
         TabIndex        =   8
         Top             =   480
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         BuddyControl    =   "txtnumdigitos"
         BuddyDispid     =   196610
         OrigLeft        =   1800
         OrigTop         =   480
         OrigRight       =   2040
         OrigBottom      =   765
         Max             =   12
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton CmdMuestra 
         Height          =   600
         Left            =   5220
         Picture         =   "FrmEstadosFinancieros.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   165
         Width           =   900
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   960
         TabIndex        =   3
         Top             =   135
         Width           =   1200
         _ExtentX        =   2117
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
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   300
         Left            =   3840
         TabIndex        =   5
         Top             =   180
         Width           =   1200
         _ExtentX        =   2117
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
      Begin VB.Label lblndigitos 
         AutoSize        =   -1  'True
         Caption         =   "Nº Digitos"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   705
      End
      Begin VB.Label lblfecfin 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Left            =   3615
         TabIndex        =   4
         Top             =   180
         Width           =   135
      End
      Begin VB.Label lblfecini 
         AutoSize        =   -1  'True
         Caption         =   "Del"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   210
         Width           =   240
      End
   End
   Begin VB.Frame Framonedas 
      Caption         =   "Monedas"
      ForeColor       =   &H00800000&
      Height          =   825
      Left            =   6240
      TabIndex        =   10
      Top             =   360
      Width           =   1455
      Begin VB.OptionButton OptSoles 
         Caption         =   "Soles"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   210
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton OptDolares 
         Caption         =   "Dolares"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13560
      Top             =   600
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
            Picture         =   "FrmEstadosFinancieros.frx":0442
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadosFinancieros.frx":0986
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadosFinancieros.frx":0D18
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadosFinancieros.frx":10AA
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadosFinancieros.frx":122E
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadosFinancieros.frx":1682
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadosFinancieros.frx":179A
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadosFinancieros.frx":1CDE
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadosFinancieros.frx":2222
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadosFinancieros.frx":2336
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadosFinancieros.frx":244A
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadosFinancieros.frx":289E
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEstadosFinancieros.frx":2A0A
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7590
      Left            =   0
      TabIndex        =   13
      Top             =   1260
      Width           =   13335
      _cx             =   23521
      _cy             =   13388
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
      Caption         =   "  &Hoja de Trabajo|   &Balance General|Estado de &Perdida y Ganacia"
      Align           =   0
      CurrTab         =   2
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
      Begin VB.Frame Frame3 
         Height          =   7170
         Left            =   45
         TabIndex        =   18
         Top             =   375
         Width           =   13245
         Begin VSFlex7Ctl.VSFlexGrid Fg3 
            Height          =   6855
            Left            =   30
            TabIndex        =   19
            Top             =   150
            Width           =   13170
            _cx             =   23230
            _cy             =   12091
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
            Rows            =   1
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmEstadosFinancieros.frx":2F52
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
      Begin VB.Frame Frame2 
         Height          =   7170
         Left            =   -13890
         TabIndex        =   16
         Top             =   375
         Width           =   13245
         Begin VSFlex7Ctl.VSFlexGrid Fg2 
            Height          =   6855
            Left            =   75
            TabIndex        =   17
            Top             =   180
            Width           =   13110
            _cx             =   23125
            _cy             =   12091
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
            Rows            =   1
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmEstadosFinancieros.frx":2FF1
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
      Begin VB.Frame FraHojaTrabajo 
         Height          =   7170
         Left            =   -14190
         TabIndex        =   14
         Top             =   375
         Width           =   13245
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   6855
            Left            =   60
            TabIndex        =   15
            Top             =   180
            Width           =   13095
            _cx             =   23098
            _cy             =   12091
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
            Rows            =   1
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmEstadosFinancieros.frx":309A
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mostrar Registro"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a Excel"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmEstadosFinancieros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CaracteresNumericos As String

Dim xMes As Integer

Sub BalanceGeneral()

Dim totala As Double
Dim totalp As Double
Dim X As Integer
Dim ACTIVO1 As Double
Dim PASIVO1 As Double
Dim ACTIVO As Double
Dim PASIVO As Double
Dim SaldoA As Double
Dim SaldoP As Double
Dim xTotalA  As Double
Dim xTotalP As Double

'No hay Problema de Tablas Vinculadas
With Fg1
    
    Fg2.AddItem ""
    
    Fg2.Col = 1
    Fg2.Row = Fg2.Rows - 1
    
    Fg2.CellFontBold = True
    
    Fg2.TextMatrix(Fg2.Rows - 1, 2) = "BALANCE GENERAL AL " & TxtFchFin.Valor
    Fg2.AddItem ""
    
    .Visible = False
    
    totala = 0
    totalp = 0
    
    For X = 1 To .Rows - 1
        .Visible = True
        If (Mid(.TextMatrix(X, 1), 1, 1) < "6") And (Mid(.TextMatrix(X, 1), 1, 1) <> "") Then
            Fg2.AddItem " "
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = .TextMatrix(X, 1)
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = .TextMatrix(X, 2)
            
            ACTIVO1 = Val(Format(.TextMatrix(X, 5), "0.00"))
            PASIVO1 = Val(Format(.TextMatrix(X, 6), "0.00"))

            ACTIVO = IIf(Val(ACTIVO1) > Val(PASIVO1), Val(ACTIVO1) - Val(PASIVO1), 0)
            PASIVO = IIf(Val(PASIVO1) > Val(ACTIVO1), Val(PASIVO1) - Val(ACTIVO1), 0)
            
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(ACTIVO, "0.00") 'Activo
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(PASIVO, "0.00") 'Pasivo
            totala = totala + ACTIVO
            totalp = totalp + PASIVO
      
        If (Mid(.TextMatrix(X, 1), 1, 1) >= "6") Then Exit For
        End If
     Next
     .Visible = True
End With

With Fg2
    .AddItem ""
    .TextMatrix(.Rows - 1, 3) = "------------------------------------------"
    .TextMatrix(.Rows - 1, 4) = "------------------------------------------"
    .AddItem " "
    .TextMatrix(.Rows - 1, 2) = "SUMATORIA DE COLUMNAS"
    .TextMatrix(.Rows - 1, 3) = Format(totala, "###,###,###,##0.00")
    .TextMatrix(.Rows - 1, 4) = Format(totalp, "###,###,###,##0.00")  'Pasivo
    SaldoA = IIf(totalp > totala, totalp - totala, 0)
    SaldoP = IIf(totala > totalp, totala - totalp, 0)
    .AddItem ""
    .TextMatrix(.Rows - 1, 2) = "RESULTADO DEL EJERCICIO"
    .TextMatrix(.Rows - 1, 3) = Format(SaldoA, "###,###,###,##0.00")
    .TextMatrix(.Rows - 1, 4) = Format(SaldoP, "###,###,###,##0.00") 'Pasivo
    xTotalA = SaldoA + totala
    xTotalP = SaldoP + totalp
    .AddItem ""
    .TextMatrix(.Rows - 1, 3) = "------------------------------------------"
    .TextMatrix(.Rows - 1, 4) = "------------------------------------------"
    .AddItem ""
    .TextMatrix(.Rows - 1, 2) = "SUMATORIA DE COLUMNAS"
    .TextMatrix(.Rows - 1, 3) = Format(xTotalA, "###,###,###,##0.00")
    .TextMatrix(.Rows - 1, 4) = Format(xTotalP, "###,###,###,##0.00") 'Pasivo
    .AddItem ""
    .TextMatrix(.Rows - 1, 3) = "------------------------------------------"
    .TextMatrix(.Rows - 1, 4) = "------------------------------------------"
End With

End Sub
Sub EstadoPerdidasGanancias()

Dim Verdadero  As Boolean
Dim Perdida1 As Double
Dim Perdida As Double
Dim Ganancia1 As Double
Dim Ganancia As Double
Dim Monto As Double
Dim totalp As Double
Dim Totalg As Double

Dim X As Integer
Dim X2 As Integer

Verdadero = False
X2 = 3
Perdida1 = 0: Ganancia1 = 0
Perdida = 0: Ganancia = 0
Monto = 0: totalp = 0
Totalg = 0
  
      
  With Fg1
    Fg3.AddItem ""
    
    Fg3.Col = 1
    Fg3.Row = Me.Fg3.Rows - 1
    Fg3.CellFontBold = True
    
    Fg3.TextMatrix(Fg3.Rows - 1, 2) = "ESTADO DE PERDIDA Y GANANCIA AL" & TxtFchFin.Valor
    Fg3.AddItem ""
    
    For X = 1 To .Rows - 1
    
    'Busca por el nro de clase
     
     Select Case Mid(.TextMatrix(X, 1), 1, 1)
      Case "6":

                If (Mid(.TextMatrix(X, 1), 1, 2) = "66") Or (Mid(.TextMatrix(X, 1), 1, 2) = "67") Or (Mid(.TextMatrix(X, 1), 1, 2) = "69") Then
                  Verdadero = True
            End If
      Case "7":
    
            If (Mid(.TextMatrix(X, 1), 1, 2) >= "70") And (Mid(.TextMatrix(X, 1), 1, 2) <= "77") Then
                    If (Mid(.TextMatrix(X, 0), 1, 2) <> "71") Then Verdadero = True
            End If
      Case "8":
            Verdadero = True
            
      Case "9":
            Verdadero = True
    
     End Select
     If (Verdadero) Then
       Fg3.AddItem " "
       Fg3.TextMatrix(X2, 1) = .TextMatrix(X, 1)
       Fg3.TextMatrix(X2, 2) = .TextMatrix(X, 2)
       Perdida = Format(Val(.TextMatrix(X, 5)), "####################0.00")
       Ganancia = Format(Val(.TextMatrix(X, 6)), "####################0.00")
       Monto = Perdida - Ganancia
       If (Monto > 0) Then
          Fg3.TextMatrix(X2, 3) = Format(Monto, "###,##0.00")
          Fg3.TextMatrix(X2, 4) = "0.00"
       End If
       If (Monto < 0) Then
          Fg3.TextMatrix(X2, 4) = Format(Abs(Monto), "###,##0.00")
          Fg3.TextMatrix(X2, 3) = "0.00"
       End If
       If (Monto = 0) Then
          Fg3.TextMatrix(X2, 3) = "0.00"
          Fg3.TextMatrix(X2, 4) = "0.00"
       End If
       Perdida = Format(Val(.TextMatrix(X, 5)), "0.00")
       Ganancia = Format(Val(.TextMatrix(X, 6)), "0.00")
       If (Monto < 0) Then
          Ganancia = IIf(Val(Ganancia) > Val(Perdida), Val(Ganancia) - Val(Perdida), 0)
          Ganancia1 = Ganancia1 + Ganancia
       Else
          Perdida = IIf(Val(Perdida) > Val(Ganancia), Val(Perdida) - Val(Ganancia), 0)
          Perdida1 = Perdida1 + Perdida
       End If
       If (Monto = 0) Then
          Ganancia = 0
          Perdida = 0
       End If
       X2 = X2 + 1
        Verdadero = False
      
      End If
    Next
  End With
  With Fg3
   .AddItem " "
     .TextMatrix(.Rows - 1, 3) = "------------------------------"
     .TextMatrix(.Rows - 1, 4) = "------------------------------"
     
   .AddItem " "
   .TextMatrix(.Rows - 1, 2) = "SUMATORIA DE COLUMNAS"
   .TextMatrix(.Rows - 1, 3) = Format(Perdida1, "###,###,###,##0.00")
   .TextMatrix(.Rows - 1, 4) = Format(Ganancia1, "###,###,###,##0.00")
   .AddItem " "
   .TextMatrix(.Rows - 1, 2) = "RESULTADO DEL EJERCICIO"
      totalp = IIf(Ganancia1 > Perdida1, Ganancia1 - Perdida1, 0)
      Totalg = IIf(Perdida1 > Ganancia1, Perdida1 - Ganancia1, 0)
   .TextMatrix(.Rows - 1, 3) = Format(totalp, "##,###,###,###,##0.00")
   .TextMatrix(.Rows - 1, 4) = Format(Totalg, "##,###,###,###,##0.00")
   .AddItem " "
     .TextMatrix(.Rows - 1, 3) = "------------------------------"
     .TextMatrix(.Rows - 1, 4) = "------------------------------"
    .AddItem " "
      totalp = totalp + Perdida1
      Totalg = Totalg + Ganancia1
   .TextMatrix(.Rows - 1, 2) = "SUMATORIA DE COLUMNAS"
   .TextMatrix(.Rows - 1, 3) = Format(totalp, "##,###,###,###,##0.00")
   .TextMatrix(.Rows - 1, 4) = Format(Totalg, "##,###,###,###,##0.00")
    .AddItem " "
     .TextMatrix(.Rows - 1, 3) = "------------------------------"
     .TextMatrix(.Rows - 1, 4) = "------------------------------"
  End With

End Sub
Sub HojaDeTrabajo()

Dim Rst As New ADODB.Recordset
Dim RSTAUX As New ADODB.Recordset
Dim TDebe As Double
Dim THaber As Double
Dim Tsaldodebe     As Double
Dim Tsaldohaber As Double


Dim numano As Integer
Dim nummes  As Integer
Dim filcta As Integer
Dim CAMPO As String
Dim GRUPO As String
Dim ORDEN As String

filcta = Val(txtnumdigitos)


CAMPO = " SELECT Left(con_planctas.cuenta, " & filcta & " ,   ) AS CuentaAgrupada, "
GRUPO = " GROUP BY Left([con_planctas]![cuenta], " & filcta & ") "
ORDEN = " ORDER BY Left([con_planctas]![cuenta], " & filcta & ") "


RST_Busq Rst, CAMPO & " Sum(con_diario.impdebsol) AS SumaDebe, Sum(con_diario.imphabsol) AS SumaHaber  " _
                 & " FROM con_planctas INNER JOIN con_diario ON con_planctas.id = con_diario.idcue " _
                 & " WHERE (((con_diario.año) = 2006) And ((con_diario.idmes) = 12)) " _
                 & GRUPO _
                 & ORDEN, xCon
                                  
 Do While Not Rst.EOF

            With Fg1
                     .AddItem ""
                     .Row = .Rows - 1
                     .TextMatrix(.Row, 1) = Rst("CuentaAgrupada")
                                                                
                      Set RSTAUX = BuscaConCriterio("SELECT con_planctas.descripcion FROM con_planctas WHERE con_planctas.cuenta ='" & Trim(Rst.Fields(0)) & "'", xCon)

                     If RSTAUX.RecordCount > 0 Then .TextMatrix(.Row, 2) = RSTAUX("Descripcion")
                     
                     .TextMatrix(.Row, 3) = NulosN(Rst("sumadebe"))
                     .TextMatrix(.Row, 4) = NulosN(Rst("sumahaber"))
                                          
                       If NulosN(Rst("sumadebe")) > NulosN(Rst("sumahaber")) Then
                            .TextMatrix(.Row, 5) = NulosN(Rst("sumadebe")) - NulosN(Rst("sumahaber"))
                                Tsaldodebe = Tsaldodebe + NulosN(Rst("sumadebe")) - NulosN(Rst("sumahaber"))
                        Else
                            .TextMatrix(.Row, 6) = NulosN(Rst("sumahaber")) - NulosN(Rst("sumadebe"))
                            Tsaldohaber = Tsaldohaber + NulosN(Rst("sumahaber")) - NulosN(Rst("sumadebe"))
                        End If
            End With
                                  TDebe = TDebe + Rst("SumaDebe")
                                  THaber = THaber + Rst("SumaHaber")
            Rst.MoveNext
 Loop
                            

                            With Fg1
                              .AddItem ""
                              .AddItem ""
                                .Row = .Rows - 1
                            .TextMatrix(.Row, 3) = "----------------------------"
                            .TextMatrix(.Row, 4) = "----------------------------"
                            .TextMatrix(.Row, 5) = "----------------------------"
                            .TextMatrix(.Row, 6) = "----------------------------"
                            
                            .AddItem vbNullString
                            .Row = .Rows - 1
                            .TextMatrix(.Row, 2) = "SUMATORIA DE COLUMNAS"
                            .TextMatrix(.Row, 3) = Format(TDebe, "########,###,###0.00")
                            .TextMatrix(.Row, 4) = Format(THaber, "########,###,###0.00")
                            .TextMatrix(.Row, 5) = Format(Tsaldodebe, "########,###,###0.00")
                            .TextMatrix(.Row, 6) = Format(Tsaldohaber, "########,###,###0.00")

                            .AddItem ""
                            .Row = .Rows - 1
                            .TextMatrix(.Row, 3) = "----------------------------"
                            .TextMatrix(.Row, 4) = "----------------------------"
                            .TextMatrix(.Row, 5) = "----------------------------"
                            .TextMatrix(.Row, 6) = "----------------------------"

       End With
                              
         Fg1.Visible = True
         'Call BalanceGeneral
         'Call EstadoPerdidas_Y_Ganancia

Set RSTAUX = Nothing
Set Rst = Nothing
End Sub

Private Sub CmdMuestra_Click()
    
    
    If NulosC(TxtFchIni.Valor) = "" Or NulosC(TxtFchFin.Valor) = "" Then
        MsgBox "El rango de fechas del periodo a consultar es invalido", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
        MsgBox "La fecha de inicio del periodo no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    If Val(txtnumdigitos.Text) <= 0 Then
      MsgBox "Ingrese Nro de Digitos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
      txtnumdigitos.SetFocus
      Exit Sub
    End If

    If Val(txtnumdigitos.Text) < 2 Or Val(Me.txtnumdigitos.Text) > 12 Then
        MsgBox "Numero de digitos entre 2..12 ", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        txtnumdigitos.SetFocus
        Exit Sub
    End If

    
    Fg1.Rows = 1
    Fg2.Rows = 1
    Fg3.Rows = 1
    Call HojaDeTrabajo
    Call BalanceGeneral
    Call EstadoPerdidasGanancias
End Sub


Private Sub Form_Load()
    Blanquea
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    txtnumdigitos = 2
    OptSoles.Value = True
    CaracteresNumericos = "0123456789" & Chr(8)
End Sub

Sub Blanquea()
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
End Sub


Private Sub OptDolares_Click()
    If OptDolares.Value = True Then
        Fg1.ColWidth(9) = 0
        Fg1.ColWidth(10) = 0
        Fg1.ColWidth(11) = 1000
        Fg1.ColWidth(12) = 1000
    End If
End Sub

Private Sub OptSoles_Click()
    If OptSoles.Value = True Then
        Fg1.ColWidth(9) = 1000
        Fg1.ColWidth(10) = 1000
        Fg1.ColWidth(11) = 0
        Fg1.ColWidth(12) = 0
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 6 Then
        Unload Me
    End If
    If Button.Index = 1 Then
        Call CmdMuestra_Click
    End If
    
    If Button.Index = 3 Then 'IMPRESION
        
    End If
    If Button.Index = 4 Then 'ENVIAR EXCEL
        
    End If
End Sub






Private Sub txtnumdigitos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  CmdMuestra.SetFocus
Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
End If


End Sub
