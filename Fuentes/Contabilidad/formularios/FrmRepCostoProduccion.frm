VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.ocx"
Begin VB.Form FrmRepCostoProduccion 
   Caption         =   "Contabilidad - Análisis de Costo de Producción"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7370
      Left            =   40
      TabIndex        =   6
      Top             =   360
      Width           =   11805
      Begin VB.CheckBox TerminadosCheck 
         Caption         =   "&Solo Productos Terminados"
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Frame Frame1 
         ForeColor       =   &H00800000&
         Height          =   1035
         Left            =   6195
         TabIndex        =   17
         Top             =   530
         Width           =   3650
         Begin VB.CommandButton cmd 
            Caption         =   "&Agregar"
            Height          =   405
            Index           =   2
            Left            =   2800
            TabIndex        =   19
            Top             =   150
            Width           =   765
         End
         Begin VB.CommandButton cmd 
            Caption         =   "&Eliminar"
            Height          =   405
            Index           =   3
            Left            =   2800
            TabIndex        =   18
            Top             =   550
            Width           =   765
         End
         Begin VSFlex7Ctl.VSFlexGrid fg 
            Height          =   795
            Index           =   2
            Left            =   45
            TabIndex        =   20
            Top             =   150
            Width           =   2730
            _cx             =   4815
            _cy             =   1402
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
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmRepCostoProduccion.frx":0000
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
         Caption         =   "[ Datos de Producción ]"
         ForeColor       =   &H00800000&
         Height          =   5785
         Left            =   0
         TabIndex        =   9
         Top             =   1580
         Width           =   11775
         Begin VB.Frame Frame7 
            Caption         =   "[ Detalles de Movimiento ]"
            ForeColor       =   &H00800000&
            Height          =   3480
            Left            =   60
            TabIndex        =   10
            Top             =   2250
            Width           =   11625
            Begin XtremeSuiteControls.PushButton ExportarButton 
               DragIcon        =   "FrmRepCostoProduccion.frx":0087
               Height          =   375
               Left            =   90
               TabIndex        =   26
               Top             =   240
               Width           =   1095
               _Version        =   786432
               _ExtentX        =   1931
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Exportar"
               Appearance      =   2
               Picture         =   "FrmRepCostoProduccion.frx":5C99
            End
            Begin VSFlex7Ctl.VSFlexGrid fg 
               Height          =   2735
               Index           =   3
               Left            =   90
               TabIndex        =   11
               Top             =   640
               Width           =   11400
               _cx             =   20108
               _cy             =   4824
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   18
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmRepCostoProduccion.frx":602B
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
         Begin VSFlex7Ctl.VSFlexGrid fg 
            Height          =   1900
            Index           =   0
            Left            =   60
            TabIndex        =   12
            Top             =   300
            Width           =   11625
            _cx             =   20505
            _cy             =   3351
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
            Cols            =   31
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmRepCostoProduccion.frx":62B3
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
      Begin VB.CommandButton cmd 
         Caption         =   "Consultar"
         Height          =   350
         Index           =   4
         Left            =   9960
         TabIndex        =   8
         ToolTipText     =   "Consultar"
         Top             =   960
         Width           =   1400
      End
      Begin VB.Frame Frame3 
         ForeColor       =   &H00800000&
         Height          =   1035
         Left            =   2520
         TabIndex        =   7
         Top             =   530
         Width           =   3650
         Begin VB.CommandButton cmd 
            Caption         =   "&Eliminar"
            Height          =   405
            Index           =   1
            Left            =   2800
            TabIndex        =   15
            Top             =   550
            Width           =   765
         End
         Begin VB.CommandButton cmd 
            Caption         =   "&Agregar"
            Height          =   405
            Index           =   0
            Left            =   2800
            TabIndex        =   14
            Top             =   150
            Width           =   765
         End
         Begin VSFlex7Ctl.VSFlexGrid fg 
            Height          =   795
            Index           =   1
            Left            =   45
            TabIndex        =   16
            Top             =   150
            Width           =   2730
            _cx             =   4815
            _cy             =   1402
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
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmRepCostoProduccion.frx":6717
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
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   870
         TabIndex        =   22
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
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
         Valor           =   "23/03/2007"
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   300
         Left            =   870
         TabIndex        =   23
         Top             =   915
         Width           =   1305
         _ExtentX        =   2302
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
         Valor           =   "23/03/2007"
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   645
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fin:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Análisis de Costo de Producción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   60
         TabIndex        =   13
         Top             =   100
         Width           =   11685
      End
   End
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   3000
      TabIndex        =   1
      Top             =   8040
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   225
         TabIndex        =   2
         Top             =   420
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cancelar = ESC"
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
         Left            =   4470
         TabIndex        =   5
         Top             =   720
         Width           =   1260
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
         Left            =   150
         TabIndex        =   4
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label LblProg 
         AutoSize        =   -1  'True
         Caption         =   "LblProg"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3500
         TabIndex        =   3
         Top             =   180
         Width           =   525
      End
      Begin VB.Shape Shape1 
         Height          =   885
         Index           =   0
         Left            =   60
         Top             =   90
         Width           =   5805
      End
   End
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
            Picture         =   "FrmRepCostoProduccion.frx":679A
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepCostoProduccion.frx":6CDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepCostoProduccion.frx":7070
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepCostoProduccion.frx":71F4
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepCostoProduccion.frx":7648
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepCostoProduccion.frx":7760
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepCostoProduccion.frx":7CA4
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepCostoProduccion.frx":81E8
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepCostoProduccion.frx":82FC
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepCostoProduccion.frx":8410
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepCostoProduccion.frx":8864
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepCostoProduccion.frx":89D0
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRepCostoProduccion.frx":8F18
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   1058
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
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
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
   Begin VB.Menu menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Insertar"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Ver Receta"
      End
   End
End
Attribute VB_Name = "FrmRepCostoProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------VARIABLES DE ESTADO DE FORMULARIO
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim Agregando As Boolean               ' para saber cuando se este agregando FILAS AL CONTROL grid de productos
Dim IdMenuActivo As Integer            ' INDICA EL CODIGO DEL MENU ACTIVO
Dim xHorIni As Date                    ' ESPECIFICA LA HORA DE INICIO
Dim mIdRegistro&                       ' identificador del registro
Dim mMesActivo As Integer              ' indica el mes activo
Dim OrigFX As Long
Dim OrigFY As Long
Dim fOrdenLista As Boolean             ' especfica el orden de la lista de la consulta
'***********************************************
'-----------------------VARIABLES DE FORMULARIO
'***********************************************
Dim RstLibro As New ADODB.Recordset
Dim cSQL As String
Dim F As New SistemaLogica.Funciones

Private Sub pLlenarDatos()
    Dim F As New SistemaLogica.Funciones
    Dim cWhere As String
    Dim mDataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    Dim mFechaInicioInventario As Date
    Dim mLCCostoProd As New ContabilidadEntidad.LECCostoProd
    
On Error GoTo BloqueError

    mFechaInicioInventario = CDate("01/01/2013")
    Agregando = True
    
    Me.MousePointer = vbHourglass
    ' Se crea el filtro de consulta
    If fg(1).Rows <> fg(1).FixedRows Then
        If cWhere <> "" Then cWhere = cWhere & " AND "
        cWhere = cWhere & GENERAR_SQL_ID(fg(1), fg(1).ColIndex("IDPARTEPROD"), "pro_producciondet.idpro")
    End If
    If fg(2).Rows <> fg(2).FixedRows Then
        If cWhere <> "" Then cWhere = cWhere & " AND "
        cWhere = cWhere & GENERAR_SQL_ID(fg(2), fg(2).ColIndex("IDORDENPROD"), "pro_producciondet.idord")
    End If
    ' Se llenan los datos
    Set mLCCostoProd.Conexion = xCon
    mLCCostoProd.Fetch TxtFchIni.Valor, TxtFchFin.Valor, mFechaInicioInventario, cWhere
    ' Se recorre la lista para calcular su importe unitario
    With fg(0)
        .Rows = .FixedRows
        fg(3).Rows = fg(3).FixedRows
        Dim mCCostoProd As New ContabilidadEntidad.ECCostoProd
        For Each mCCostoProd In mLCCostoProd
            If TerminadosCheck.Value = 1 And mCCostoProd.TipoItem <> "PROT" Then GoTo Saltar_Siguiente
            
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("FECHA")) = mCCostoProd.Fecha
            .TextMatrix(.Rows - 1, .ColIndex("TIPMOV")) = mCCostoProd.TipoMovimiento
            .TextMatrix(.Rows - 1, .ColIndex("NUMMOV")) = mCCostoProd.MovimientoAlmacen
            .TextMatrix(.Rows - 1, .ColIndex("ALMACEN")) = mCCostoProd.Almacen
            .TextMatrix(.Rows - 1, .ColIndex("ORDENPROD")) = mCCostoProd.OrdenProduccion
            .TextMatrix(.Rows - 1, .ColIndex("PARTEPROD")) = mCCostoProd.ParteProduccion
            .TextMatrix(.Rows - 1, .ColIndex("TIPO")) = mCCostoProd.TipoItem
            .TextMatrix(.Rows - 1, .ColIndex("CODITEM")) = mCCostoProd.CodigoItem
            .TextMatrix(.Rows - 1, .ColIndex("ITEM")) = mCCostoProd.Item
            .TextMatrix(.Rows - 1, .ColIndex("RECETA")) = mCCostoProd.Receta
            .TextMatrix(.Rows - 1, .ColIndex("UM")) = mCCostoProd.UnidadMedida
            .TextMatrix(.Rows - 1, .ColIndex("CANTIDAD")) = Format(mCCostoProd.Cantidad, FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, .ColIndex("HORINI")) = Format(F.NuloString(mCCostoProd.HoraInicio), "HH:mm")
            .TextMatrix(.Rows - 1, .ColIndex("HORFIN")) = Format(F.NuloString(mCCostoProd.HoraFin), "HH:mm")
            .TextMatrix(.Rows - 1, .ColIndex("COSTOMP")) = Format(mCCostoProd.CostoMP, FORMAT_IMPORTEKARDEX)
            .TextMatrix(.Rows - 1, .ColIndex("COSTOMOD")) = Format(mCCostoProd.CostoMOD, FORMAT_IMPORTEKARDEX)
            .TextMatrix(.Rows - 1, .ColIndex("COSTOPRIMO")) = Format(mCCostoProd.CostoPrimo, FORMAT_IMPORTEKARDEX)
            .TextMatrix(.Rows - 1, .ColIndex("COSTOCIF")) = Format(mCCostoProd.CostoCIF, FORMAT_IMPORTEKARDEX)
            .TextMatrix(.Rows - 1, .ColIndex("COSTOTOTAL")) = Format(mCCostoProd.CostoTotal, FORMAT_IMPORTEKARDEX)
            .TextMatrix(.Rows - 1, .ColIndex("CUPROD")) = Format(mCCostoProd.CostoUP, FORMAT_IMPORTEKARDEX)
            .TextMatrix(.Rows - 1, .ColIndex("PVENTA")) = Format(mCCostoProd.PrecioVenta, FORMAT_IMPORTEKARDEX)
            .TextMatrix(.Rows - 1, .ColIndex("IMPVENTA")) = Format(mCCostoProd.ImporteVenta, FORMAT_IMPORTEKARDEX)
            .TextMatrix(.Rows - 1, .ColIndex("DESV")) = Format(mCCostoProd.Desviacion, FORMAT_IMPORTEKARDEX)
            .TextMatrix(.Rows - 1, .ColIndex("DESVPORC")) = Format(mCCostoProd.DesviacionPorc, FORMAT_PORCENTAJE)
            .TextMatrix(.Rows - 1, .ColIndex("IDPARTEPROD")) = mCCostoProd.IdParteProduccion
            .TextMatrix(.Rows - 1, .ColIndex("IDPARTEPRODDET")) = mCCostoProd.IdParteProduccionDet
            .TextMatrix(.Rows - 1, .ColIndex("IDMOVDET")) = mCCostoProd.IdMovimientoDetalle
            .TextMatrix(.Rows - 1, .ColIndex("IDITEM")) = mCCostoProd.IdItem
            
Saltar_Siguiente:
        Next
        ' Se agrega la fila de totales
        .Rows = .Rows + 1
        FORMATO_CELDA fg(0), .Rows - 1, .ColIndex("ITEM"), , True, , "TOTAL"
        .TextMatrix(.Rows - 1, .ColIndex("COSTOMP")) = Format(GRID_SUMAR_COL(fg(0), .ColIndex("COSTOMP")), FORMAT_IMPORTEKARDEX)
        .TextMatrix(.Rows - 1, .ColIndex("COSTOMOD")) = Format(GRID_SUMAR_COL(fg(0), .ColIndex("COSTOMOD")), FORMAT_IMPORTEKARDEX)
        .TextMatrix(.Rows - 1, .ColIndex("COSTOPRIMO")) = Format(GRID_SUMAR_COL(fg(0), .ColIndex("COSTOPRIMO")), FORMAT_IMPORTEKARDEX)
        .TextMatrix(.Rows - 1, .ColIndex("COSTOCIF")) = Format(GRID_SUMAR_COL(fg(0), .ColIndex("COSTOCIF")), FORMAT_IMPORTEKARDEX)
        .TextMatrix(.Rows - 1, .ColIndex("COSTOTOTAL")) = Format(GRID_SUMAR_COL(fg(0), .ColIndex("COSTOTOTAL")), FORMAT_IMPORTEKARDEX)
        .TextMatrix(.Rows - 1, .ColIndex("CUPROD")) = Format(GRID_SUMAR_COL(fg(0), .ColIndex("CUPROD")), FORMAT_IMPORTEKARDEX)
        .TextMatrix(.Rows - 1, .ColIndex("PVENTA")) = Format(GRID_SUMAR_COL(fg(0), .ColIndex("PVENTA")), FORMAT_IMPORTEKARDEX)
        .TextMatrix(.Rows - 1, .ColIndex("IMPVENTA")) = Format(GRID_SUMAR_COL(fg(0), .ColIndex("IMPVENTA")), FORMAT_IMPORTEKARDEX)
        .TextMatrix(.Rows - 1, .ColIndex("DESV")) = Format(GRID_SUMAR_COL(fg(0), .ColIndex("DESV")), FORMAT_IMPORTEKARDEX)
        .TopRow = .Rows - 1
    End With
    
    Me.MousePointer = vbDefault
    Agregando = False
    Exit Sub
    
BloqueError:
    Agregando = False
    Me.MousePointer = vbDefault
    F.MostrarMensajeError Err.Description, "LlenarDatos", Err.Source, Err.Number
End Sub

Private Sub llenarDetalleInsumos()
    Dim mImporteTeorico As Double
    Dim mImporteMOD As Double
    Dim mImporteCIF As Double
    Dim mCantidadCabecera As Double
    
    If Agregando Then Exit Sub
    If fg(0).Rows = fg(0).FixedRows Then Exit Sub
    
    fg(3).Rows = fg(3).FixedRows
    With fg(3)
        ' Obtenemos el Parte detallado seleccionado
        Dim mParteProdDet As New ProduccionEntidad.EParteProdDet
        mParteProdDet.LoadChild = True
        Set mParteProdDet.Conexion = xCon
        mParteProdDet.Fetch F.NuloNumeric(fg(0).TextMatrix(fg(0).Row, fg(0).ColIndex("IDPARTEPRODDET")))
        
        Dim mParteProdDetIns As New ProduccionEntidad.EParteProdDetIns
        For Each mParteProdDetIns In mParteProdDet.LParteProduccionDetIns
            Dim mParteProdDetInsMov As New ProduccionEntidad.EParteProdDetInsMov
            For Each mParteProdDetInsMov In mParteProdDetIns.LParteProduccionDetInsMov
                Dim mFecha As String
                Dim mIdMovimiento As Long
                Dim mNumMov As String
                Dim mTipoMovimiento As Long
                Dim mIdAlmacen As Long
                Dim mAlmacen As String
                Dim mIdTipDocRef As Long
                Dim mTipDocRef As String
                Dim mIdDocRef As Long
                Dim mNumDocRef As String
                Dim mCostoUnitarioPromedio As Double
                Dim mImporte As Double
                Dim mDataBase As New SistemaData.EDataBase
                Dim mRecord As New ADODB.Recordset
                
                mIdMovimiento = F.NuloNumeric(F.BuscaCodigoTabla(mParteProdDetInsMov.IdMovimientoDetalle, "idmovdet", "id", "alm_ingresodet", "N", xCon))
                
                mFecha = F.NuloString(F.BuscaCodigoTabla(mIdMovimiento, "id", "fchdoc", "alm_ingreso", "N", xCon))
                mNumMov = F.NuloString(F.BuscaCodigoTabla(mIdMovimiento, "id", "numser", "alm_ingreso", "N", xCon))
                mNumMov = mNumMov & "-" & F.NuloString(F.BuscaCodigoTabla(mIdMovimiento, "id", "numdoc", "alm_ingreso", "N", xCon))
                mIdAlmacen = F.NuloNumeric(F.BuscaCodigoTabla(mIdMovimiento, "id", "idalm", "alm_ingreso", "N", xCon))
                mAlmacen = F.NuloString(F.BuscaCodigoTabla(mIdAlmacen, "id", "descripcion", "alm_almacenes", "N", xCon))
                mTipoMovimiento = F.NuloNumeric(F.BuscaCodigoTabla(mIdMovimiento, "id", "tipmov", "alm_ingreso", "N", xCon))
                mIdTipDocRef = F.NuloNumeric(F.BuscaCodigoTabla(mIdMovimiento, "id", "idtipdocref", "alm_ingreso", "N", xCon))
                mTipDocRef = F.NuloString(F.BuscaCodigoTabla(mIdTipDocRef, "id", "abrev", "mae_documento", "N", xCon))
                If mIdTipDocRef = F.NuloNumeric(F.KeyValue("IdDocumentoSolictudMateriales", xCon)) Then
                    mIdDocRef = F.NuloNumeric(F.BuscaCodigoTabla(mIdMovimiento, "id", "iddocref", "alm_ingreso", "N", xCon))
                    mNumDocRef = F.NuloString(F.BuscaCodigoTabla(mIdDocRef, "id", "numser", "pro_solicitudmat", "N", xCon))
                    mNumDocRef = mNumDocRef & "-" & F.NuloString(F.BuscaCodigoTabla(mIdDocRef, "id", "numdoc", "pro_solicitudmat", "N", xCon))
                End If
                mCostoUnitarioPromedio = F.NuloNumeric(F.BuscaCodigoTabla(mParteProdDetInsMov.IdMovimientoDetalle, "idmovdet", "costounitariopromedio", "con_librocostotemp", "N", xCon))
                mImporte = F.NuloNumeric(F.BuscaCodigoTabla(mParteProdDetInsMov.IdMovimientoDetalle, "idmovdet", "costoprimo", "con_librocostotemp", "N", xCon))
                
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("FECHA")) = mFecha
                .TextMatrix(.Rows - 1, .ColIndex("CODIGOITEM")) = mParteProdDetIns.CodigoItem
                .TextMatrix(.Rows - 1, .ColIndex("ITEM")) = mParteProdDetIns.Item
                .TextMatrix(.Rows - 1, .ColIndex("NUMMOV")) = mNumMov
                .TextMatrix(.Rows - 1, .ColIndex("ALMACEN")) = mAlmacen
                .TextMatrix(.Rows - 1, .ColIndex("TIPDOCREF")) = mTipDocRef
                .TextMatrix(.Rows - 1, .ColIndex("NUMDOCREF")) = mNumDocRef
                If mTipoMovimiento = 0 Then
                    .TextMatrix(.Rows - 1, .ColIndex("TIPOMOV")) = "S"
                Else
                    .TextMatrix(.Rows - 1, .ColIndex("TIPOMOV")) = "I"
                End If
                .TextMatrix(.Rows - 1, .ColIndex("CANTIDAD")) = Format(mParteProdDetInsMov.Cantidad, FORMAT_CANTIDAD)
                .TextMatrix(.Rows - 1, .ColIndex("COSTOPROMEDIO")) = Format(mCostoUnitarioPromedio, FORMAT_IMPORTEKARDEX)
                .TextMatrix(.Rows - 1, .ColIndex("IMPORTE")) = Format(mImporte, FORMAT_IMPORTEKARDEX)
                
                Set mDataBase.Connection = xCon
                mDataBase.CommandText = "SELECT alm_inventario.codpro, alm_inventario.tippro AS idtippro, pro_recetains.iditem, [pro_recetains]![canpro]*" & mParteProdDet.CantidadProducida & " AS cantidad, pro_recetains.idunimed " _
                        + vbCr + "FROM pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id " _
                        + vbCr + "WHERE (((pro_recetains.idrec)=" & mParteProdDet.IdReceta & ") AND (alm_inventario.id) = " & mParteProdDetIns.IdItem & ");"
                Set mRecord = mDataBase.GetRecordset
                If mRecord.RecordCount > 0 Then
                    .TextMatrix(.Rows - 1, .ColIndex("CANTEO")) = Format(F.NuloNumeric(mRecord("cantidad")), FORMAT_CANTIDAD)
                    .TextMatrix(.Rows - 1, .ColIndex("COSTEO")) = Format(mCostoUnitarioPromedio * F.NuloNumeric(mRecord("cantidad")), FORMAT_IMPORTEKARDEX)
                    .TextMatrix(.Rows - 1, .ColIndex("VARCANT")) = Format(F.NuloNumeric(mRecord("cantidad")) - mParteProdDetInsMov.Cantidad, FORMAT_CANTIDAD)
                    .TextMatrix(.Rows - 1, .ColIndex("VARCOSTO")) = Format(F.NuloNumeric(.TextMatrix(.Rows - 1, .ColIndex("COSTEO"))) - mImporte, FORMAT_IMPORTEKARDEX)
                End If
                Set mRecord = Nothing
            Next
        Next
        ' Se agrega la fila de totales
        .Rows = .Rows + 1
        FORMATO_CELDA fg(3), .Rows - 1, .ColIndex("ITEM"), , True, , "TOTAL"
        .TextMatrix(.Rows - 1, .ColIndex("IMPORTE")) = Format(GRID_SUMAR_COL(fg(3), .ColIndex("IMPORTE")), FORMAT_IMPORTEKARDEX)
        mImporteTeorico = GRID_SUMAR_COL(fg(3), .ColIndex("COSTEO"))
        .TextMatrix(.Rows - 1, .ColIndex("COSTEO")) = Format(mImporteTeorico, FORMAT_IMPORTEKARDEX)
        .TextMatrix(.Rows - 1, .ColIndex("VARCOSTO")) = Format(GRID_SUMAR_COL(fg(3), .ColIndex("VARCOSTO")), FORMAT_IMPORTEKARDEX)
        
        ' Se agrega la fila de CUP Teorico
        .Rows = .Rows + 1
        FORMATO_CELDA fg(3), .Rows - 1, .ColIndex("ITEM"), , True, , "CUP TEÓRICO"
        mImporteMOD = F.NuloNumeric(fg(0).TextMatrix(fg(0).Row, fg(0).ColIndex("COSTOMOD")))
        mImporteCIF = F.NuloNumeric(fg(0).TextMatrix(fg(0).Row, fg(0).ColIndex("COSTOCIF")))
        mCantidadCabecera = F.NuloNumeric(fg(0).TextMatrix(fg(0).Row, fg(0).ColIndex("CANTIDAD")))
        .TextMatrix(.Rows - 1, .ColIndex("COSTEO")) = Format((mImporteTeorico + mImporteCIF + mImporteMOD) / mCantidadCabecera, FORMAT_IMPORTEKARDEX)
        
        .TopRow = .Rows - 1
    End With
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim A As Integer
    Dim num As Integer
    Dim Rpta As Integer
    Dim nTitulo As String
    Dim xRs As New ADODB.Recordset
    Dim xCampos() As String
    Dim xform As New eps_librerias.FormSeleccion
    Dim MENSAJE_ As String
    Dim nSQLId As String
    Dim nSQLId2 As String
    Dim NUMEROMAXTRAB_ As Integer
    Dim NUMREGAAGREGAR_ As Integer
    
    Dim ULTIMODIAMES_ As Date
    Dim PRIMERDIAMES_ As Date
    Dim ANIOACTUAL_ As Integer
    Dim MESACTUAL_ As Integer
        
            
    Select Case Index
        Case 0 ' Agregar Parte
            ReDim xCampos(2, 4) As String

            xCampos(0, 0) = "Fecha":            xCampos(0, 1) = "fchdoc":           xCampos(0, 2) = "1200":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Numero Parte":     xCampos(1, 1) = "numparteprod":     xCampos(1, 2) = "2300":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
            
            nSQLId = GENERAR_SQL_ID(fg(1), fg(1).ColIndex("IDPARTEPROD"), " AND pro_produccion.id", "NOT IN")
            cSQL = "SELECT pro_produccion.id, pro_produccion.fchdoc, [pro_produccion].[numser] & '-' & [pro_produccion].[numdoc] AS numparteprod " _
                + vbCr + "FROM pro_produccion " _
                + vbCr + "WHERE (((pro_produccion.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (pro_produccion.fchdoc)<=CDate('" & TxtFchFin.Valor & "'))) " & nSQLId

            nTitulo = "Buscando Partes de Producción"
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "fchdoc", "numparteprod", CualquierParte

            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            With fg(1)
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("PARTEPROD")) = F.NuloString(xRs("numparteprod"))
                .TextMatrix(.Rows - 1, .ColIndex("IDPARTEPROD")) = F.NuloNumeric(xRs("id"))
                .Select .Rows - 1, 1
                .TopRow = .Rows - 1
            End With
            
        Case 1 ' Eliminar Parte
            If fg(1).Rows = fg(1).FixedRows Then Exit Sub
            Rpta = MsgBox("¿Esta seguro de quitar el registro seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then fg(1).RemoveItem fg(1).Row
            
        Case 2 ' Agregar Orden de Produccion
            ReDim xCampos(2, 4) As String

            xCampos(0, 0) = "Fecha":            xCampos(0, 1) = "fchpro":       xCampos(0, 2) = "1200":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Numero Orden":     xCampos(1, 1) = "ordenprod":    xCampos(1, 2) = "2300":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
            
            nSQLId = GENERAR_SQL_ID(fg(2), fg(2).ColIndex("IDORDENPROD"), " AND pro_ordenprod.id", "NOT IN")
            cSQL = "SELECT pro_ordenprod.id, pro_ordenprod.fchpro, [pro_ordenprod].[numser] & '-' & [pro_ordenprod].[numdoc] AS ordenprod " _
                + vbCr + "FROM pro_ordenprod " _
                + vbCr + "WHERE (((pro_ordenprod.fchpro)<=CDate('" & TxtFchFin.Valor & "')) AND ((pro_ordenprod.estado)=2)) " & nSQLId

            nTitulo = "Buscando Partes de Producción"
            Set xRs = Nothing
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "fchpro", "ordenprod", CualquierParte

            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            With fg(2)
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("ORDENPROD")) = F.NuloString(xRs("ordenprod"))
                .TextMatrix(.Rows - 1, .ColIndex("IDORDENPROD")) = F.NuloNumeric(xRs("id"))
                .Select .Rows - 1, 1
                .TopRow = .Rows - 1
            End With
            
        Case 3 ' Eliminar Orden de Produccion
            If fg(2).Rows = fg(2).FixedRows Then Exit Sub
            Rpta = MsgBox("¿Esta seguro de quitar el registro seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then fg(2).RemoveItem fg(2).Row
                    
        Case 4 ' Consultar
            pLlenarDatos
            
    End Select
End Sub

Private Sub ExportarButton_Click()
    ExportarExcel fg(3), "Detalle de Costos de Producción"
End Sub

Private Sub fg_DblClick(Index As Integer)
    If Index <> 0 Then Exit Sub
    If Agregando Then Exit Sub
    If (fg(0).Row = fg(0).Rows - 1) And QueHace = 3 Then Exit Sub
    
    Me.MousePointer = vbHourglass
    llenarDetalleInsumos
    Me.MousePointer = vbDefault
End Sub

Private Sub fg_EnterCell(Index As Integer)
    fg(Index).Editable = flexEDNone
End Sub

Private Sub Form_Load()
    QueHace = 3
    SeEjecuto = False
    Agregando = False
    iniciarCampos
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDOS E CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        mMesActivo = xMes
            
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        '--interrumpir
        'BANDERA_ = True
    End If
End Sub

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub

    If Me.Width <= 12000 Then Me.Width = 12000
    If Me.Height <= 4000 Then Me.Height = 4000
        
    ' Se dimensiona el Detalle
    Label5.Width = Me.Width - 100
    
    Frame2.Width = Me.Width - 258
    Frame2.Height = Me.Height - 955
    
    Frame4.Width = Me.Width - 315
    Frame4.Height = Me.Height - 2540
    
    fg(0).Width = Frame4.Width - 150
    fg(0).Height = Frame4.Height - 3885
    
    Frame7.Top = Frame4.Height - 3535
    Frame7.Width = Frame4.Width - 150
    fg(3).Width = Frame7.Width - 225
End Sub

Private Sub iniciarCampos()
    '**********************
    ' CONFIGURACIONES GRID
    '**********************
    fg(0).AllowUserResizing = flexResizeColumns
    fg(0).AutoSearch = flexSearchFromTop
    fg(0).ExplorerBar = flexExSortShow
    fg(0).SelectionMode = flexSelectionByRow
    fg(0).Editable = flexEDNone
    fg(0).ForeColorSel = &H80000005
    fg(0).BackColorSel = &H80&
    fg(0).Rows = fg(0).FixedRows
    fg(0).FrozenCols = fg(0).ColIndex("ITEM")
    
    
    fg(1).AllowUserResizing = flexResizeColumns
    fg(1).AutoSearch = flexSearchFromTop
    fg(1).ExplorerBar = flexExSortShow
    fg(1).SelectionMode = flexSelectionByRow
    fg(1).Editable = flexEDNone
    fg(1).ForeColorSel = &H80000005
    fg(1).BackColorSel = &H80&
    fg(1).Rows = fg(1).FixedRows
    
    fg(2).AllowUserResizing = flexResizeColumns
    fg(2).AutoSearch = flexSearchFromTop
    fg(2).ExplorerBar = flexExSortShow
    fg(2).SelectionMode = flexSelectionByRow
    fg(2).Editable = flexEDNone
    fg(2).ForeColorSel = &H80000005
    fg(2).BackColorSel = &H80&
    fg(2).Rows = fg(2).FixedRows
    
    fg(3).AllowUserResizing = flexResizeColumns
    fg(3).AutoSearch = flexSearchFromTop
    fg(3).ExplorerBar = flexExSortShow
    fg(3).SelectionMode = flexSelectionByRow
    fg(3).ForeColorSel = &H80000005
    fg(3).BackColorSel = &H80&
    fg(3).Editable = flexEDKbdMouse
    fg(3).Rows = fg(3).FixedRows
        
    TxtFchIni.Valor = CDate("01/" & Month(Date) & "/" & Year(Date) & "") 'Date
    TxtFchFin.Valor = CDate("01/" & Month(Date) & "/" & Year(Date) & "") 'Date
    TerminadosCheck.Value = 1
End Sub

Sub ActivaTool()
    Dim A As Integer
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 10 Then pLlenarDatos
        
    If Button.Index = 14 Then ExportarExcel fg(0), "Costos de Producción"
    
    If Button.Index = 17 Then Unload Me
End Sub

Sub ExportarExcel(ByRef Grid As VSFlexGrid, Titulo As String)
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    Dim TITULO_ As String
            
    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, Grid, Titulo, "", "", Titulo
    Set X_EXPORT = Nothing
    MousePointer = vbDefault
    Exit Sub
error:
    MousePointer = vbDefault
    SHOW_ERROR Name, "Exportar"
End Sub
