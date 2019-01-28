VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmHojaTrabajo1 
   Caption         =   "Contabilidad - Hoja de Trabajo"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   13215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDetalle 
      BorderStyle     =   0  'None
      Height          =   6000
      Left            =   12600
      TabIndex        =   38
      Top             =   1620
      Visible         =   0   'False
      Width           =   11745
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   11490
         Picture         =   "FrmHojaTrabajo1.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   39
         ToolTipText     =   "Cerrar"
         Top             =   60
         Width           =   195
      End
      Begin VSFlex7Ctl.VSFlexGrid fg2 
         Height          =   5235
         Left            =   75
         TabIndex        =   40
         Top             =   660
         Width           =   11580
         _cx             =   20426
         _cy             =   9234
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmHojaTrabajo1.frx":02EC
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
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   345
         Left            =   90
         TabIndex        =   42
         Top             =   300
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   609
         ButtonWidth     =   609
         ButtonHeight    =   556
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Consultar"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Exportar a MSExcel"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Salir"
               ImageIndex      =   11
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label lblCuenta 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "DETALLE DE LA CUENTA"
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
         Left            =   60
         TabIndex        =   41
         Top             =   60
         Width           =   2250
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   11730
         X2              =   11730
         Y1              =   0
         Y2              =   10335
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   11700
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   0
         X1              =   30
         X2              =   11700
         Y1              =   5970
         Y2              =   5970
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   0
         X1              =   30
         X2              =   30
         Y1              =   -45
         Y2              =   5970
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   255
         Index           =   0
         Left            =   30
         Top             =   30
         Width           =   11685
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6120
      Left            =   30
      TabIndex        =   48
      Top             =   1500
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
      Caption         =   "   Por Detalle   |    Por Cuenta    |  Por Sub Cuenta  |   Análisis  "
      Align           =   0
      CurrTab         =   3
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
      Flags(2)        =   2
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         Height          =   5760
         Left            =   15
         TabIndex        =   53
         Top             =   15
         Width           =   11820
         Begin VB.Frame Frame13 
            Caption         =   "[ Cuadro de opciones a mostrar ]"
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
            Height          =   690
            Left            =   30
            TabIndex        =   55
            Top             =   30
            Width           =   11760
            Begin VB.CheckBox chk 
               Caption         =   "GF: Ganancias por Función"
               Height          =   195
               Index           =   7
               Left            =   8940
               TabIndex        =   63
               Top             =   435
               Value           =   1  'Checked
               Width           =   2490
            End
            Begin VB.CheckBox chk 
               Caption         =   "GN: Ganancias por Naturaleza"
               Height          =   210
               Index           =   6
               Left            =   8940
               TabIndex        =   62
               Top             =   210
               Value           =   1  'Checked
               Width           =   2580
            End
            Begin VB.CheckBox chk 
               Caption         =   "CT: Cuentas de Transferencia"
               Height          =   195
               Index           =   5
               Left            =   5895
               TabIndex        =   61
               Top             =   435
               Value           =   1  'Checked
               Width           =   2565
            End
            Begin VB.CheckBox chk 
               Caption         =   "CB: Cuentas de Balance"
               Height          =   210
               Index           =   4
               Left            =   5895
               TabIndex        =   60
               Top             =   210
               Value           =   1  'Checked
               Width           =   2370
            End
            Begin VB.CheckBox chk 
               Caption         =   "SF: Saldos Finales"
               Height          =   195
               Index           =   3
               Left            =   3030
               TabIndex        =   59
               Top             =   435
               Value           =   1  'Checked
               Width           =   2085
            End
            Begin VB.CheckBox chk 
               Caption         =   "SM: Sumas del Mayor"
               Height          =   210
               Index           =   2
               Left            =   3030
               TabIndex        =   58
               Top             =   210
               Value           =   1  'Checked
               Width           =   1860
            End
            Begin VB.CheckBox chk 
               Caption         =   "MP: Movimientos del Periodo"
               Height          =   195
               Index           =   1
               Left            =   165
               TabIndex        =   57
               Top             =   435
               Value           =   1  'Checked
               Width           =   2415
            End
            Begin VB.CheckBox chk 
               Caption         =   "SI: Saldos Iniciales"
               Height          =   210
               Index           =   0
               Left            =   165
               TabIndex        =   56
               Top             =   210
               Value           =   1  'Checked
               Width           =   1860
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid fg1 
            Height          =   4980
            Index           =   3
            Left            =   30
            TabIndex        =   54
            Top             =   750
            Width           =   11760
            _cx             =   20743
            _cy             =   8784
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
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   22
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmHojaTrabajo1.frx":048C
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
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   5760
         Left            =   -12735
         TabIndex        =   50
         Top             =   15
         Width           =   11820
         Begin VB.Frame Frame14 
            Caption         =   "Cuadro de Opciones"
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
            Height          =   690
            Left            =   30
            TabIndex        =   64
            Top             =   30
            Width           =   11745
            Begin VB.OptionButton Opt2Digito 
               Caption         =   "2 Dígitos"
               Height          =   195
               Left            =   300
               TabIndex        =   68
               Tag             =   "2"
               Top             =   300
               Value           =   -1  'True
               Width           =   1155
            End
            Begin VB.OptionButton Opt3Digito 
               Caption         =   "3 Dígitos"
               Height          =   195
               Left            =   1590
               TabIndex        =   67
               Tag             =   "3"
               Top             =   330
               Width           =   1155
            End
            Begin VB.CommandButton CmdExpPDT 
               Height          =   540
               Left            =   11130
               Picture         =   "FrmHojaTrabajo1.frx":0690
               Style           =   1  'Graphical
               TabIndex        =   66
               ToolTipText     =   "Exportar a PDT"
               Top             =   120
               Width           =   585
            End
            Begin VB.CheckBox Chk0Decimales 
               Caption         =   "0 Decimales"
               Height          =   225
               Left            =   3930
               TabIndex        =   65
               Top             =   330
               Width           =   1245
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid fg1 
            Height          =   4980
            Index           =   1
            Left            =   30
            TabIndex        =   51
            Top             =   750
            Width           =   11745
            _cx             =   20717
            _cy             =   8784
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
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   22
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmHojaTrabajo1.frx":1152
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
      Begin VSFlex7Ctl.VSFlexGrid fg1 
         Height          =   5760
         Index           =   0
         Left            =   -13035
         TabIndex        =   49
         Top             =   15
         Width           =   11820
         _cx             =   20849
         _cy             =   10160
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   22
         FixedRows       =   2
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmHojaTrabajo1.frx":1356
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
      Begin VSFlex7Ctl.VSFlexGrid fg1 
         Height          =   5760
         Index           =   2
         Left            =   -12435
         TabIndex        =   52
         Top             =   15
         Width           =   11820
         _cx             =   20849
         _cy             =   10160
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   22
         FixedRows       =   2
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmHojaTrabajo1.frx":156A
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
   Begin VB.Frame fra 
      Caption         =   "[ Tipo de Personal]"
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
      Height          =   585
      Index           =   4
      Left            =   3420
      TabIndex        =   29
      Top             =   885
      Width           =   8385
      Begin VB.CommandButton cb 
         Height          =   225
         Index           =   0
         Left            =   435
         Picture         =   "FrmHojaTrabajo1.frx":176E
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   255
         Width           =   210
      End
      Begin VB.CommandButton cb 
         Height          =   225
         Index           =   1
         Left            =   3325
         Picture         =   "FrmHojaTrabajo1.frx":18A0
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   255
         Width           =   210
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   1
         Left            =   2055
         MaxLength       =   15
         TabIndex        =   31
         Tag             =   "null"
         Text            =   "txt_cb(1)"
         Top             =   225
         Width           =   1515
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   0
         Left            =   30
         MaxLength       =   20
         TabIndex        =   35
         Text            =   "txt_cb(0)"
         Top             =   225
         Width           =   645
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
         Height          =   285
         Index           =   0
         Left            =   1035
         TabIndex        =   36
         Top             =   210
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl_cod 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cod(1)"
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
         Height          =   285
         Index           =   1
         Left            =   5265
         TabIndex        =   33
         Top             =   255
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl_cb 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cb(1)"
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
         Height          =   270
         Index           =   1
         Left            =   3570
         TabIndex        =   32
         Top             =   225
         Width           =   4740
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
         Height          =   270
         Index           =   0
         Left            =   690
         TabIndex        =   37
         Top             =   225
         Width           =   1215
      End
   End
   Begin VB.Frame Frame11 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   705
      Left            =   1560
      TabIndex        =   43
      Top             =   7710
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   45
         TabIndex        =   44
         Top             =   330
         Width           =   5820
         _ExtentX        =   10266
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Procesando Consulta"
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
         Left            =   60
         TabIndex        =   46
         Top             =   90
         Width           =   1770
      End
      Begin VB.Label lbl 
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
         Left            =   4365
         TabIndex        =   45
         Top             =   90
         Width           =   1530
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   2
         X1              =   -30
         X2              =   5940
         Y1              =   0
         Y2              =   15
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   5925
         X2              =   5925
         Y1              =   15
         Y2              =   945
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   5925
         Y1              =   690
         Y2              =   675
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "[ Expresado en ]"
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
      Height          =   555
      Left            =   3420
      TabIndex        =   17
      Top             =   345
      Width           =   3075
      Begin VB.CommandButton CmdBusMon 
         Height          =   230
         Left            =   1030
         Picture         =   "FrmHojaTrabajo1.frx":19D2
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   225
         Width           =   210
      End
      Begin VB.TextBox TxtIdMon 
         Height          =   300
         Left            =   720
         MaxLength       =   1
         TabIndex        =   4
         Text            =   "TxtIdMon"
         Top             =   195
         Width           =   555
      End
      Begin VB.Label LblMoneda 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblMoneda"
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
         Height          =   285
         Left            =   1260
         TabIndex        =   20
         Top             =   210
         Width           =   1755
      End
      Begin VB.Label LblTipCam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda"
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   19
         Top             =   270
         Width           =   585
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Seleccionar Libro"
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
      Height          =   555
      Left            =   6510
      TabIndex        =   24
      Top             =   345
      Width           =   5295
      Begin VB.CheckBox ChkLibro 
         Height          =   195
         Left            =   1650
         TabIndex        =   28
         Top             =   0
         Width           =   165
      End
      Begin VB.CommandButton CmdBusProv 
         Enabled         =   0   'False
         Height          =   225
         Left            =   4960
         Picture         =   "FrmHojaTrabajo1.frx":1B04
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   210
      End
      Begin VB.TextBox TxtLibro 
         Enabled         =   0   'False
         Height          =   300
         Left            =   465
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "TxtLibro"
         Top             =   195
         Width           =   4740
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Libro"
         Height          =   195
         Left            =   60
         TabIndex        =   27
         Top             =   300
         Width           =   345
      End
      Begin VB.Label LblIdLibro 
         AutoSize        =   -1  'True
         Caption         =   "LblIdLibro"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   630
         TabIndex        =   26
         Top             =   570
         Visible         =   0   'False
         Width           =   690
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "[ Consulta ]"
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
      Height          =   1125
      Left            =   30
      TabIndex        =   21
      Top             =   345
      Width           =   1515
      Begin VB.OptionButton opt_fecha 
         Caption         =   "Por Fecha"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton opt_fecha 
         Caption         =   "Por Periodo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   22
         Top             =   750
         Width           =   1125
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "[ Seleccionar ]"
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
      Height          =   1125
      Left            =   1560
      TabIndex        =   9
      Top             =   345
      Width           =   1845
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   765
         Left            =   30
         TabIndex        =   12
         Top             =   195
         Visible         =   0   'False
         Width           =   1740
         Begin VB.CommandButton cmd_periodo1 
            Height          =   240
            Left            =   1380
            Picture         =   "FrmHojaTrabajo1.frx":1C36
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   60
            Width           =   270
         End
         Begin VB.CommandButton cmd_periodo2 
            Height          =   240
            Left            =   1380
            Picture         =   "FrmHojaTrabajo1.frx":1FB8
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   450
            Width           =   270
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "A"
            Height          =   195
            Left            =   30
            TabIndex        =   16
            Top             =   480
            Width           =   105
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "De"
            Height          =   195
            Left            =   30
            TabIndex        =   15
            Top             =   120
            Width           =   210
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
            Left            =   330
            TabIndex        =   14
            Top             =   420
            Width           =   1365
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
            Left            =   330
            TabIndex        =   13
            Top             =   30
            Width           =   1365
         End
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   420
         TabIndex        =   0
         Top             =   270
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
         Left            =   420
         TabIndex        =   3
         Top             =   675
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Del"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Left            =   75
         TabIndex        =   10
         Top             =   765
         Width           =   135
      End
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   60
      TabIndex        =   6
      Top             =   9000
      Width           =   11865
      Begin VB.Label LblDescCuenta 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblDescCuenta"
         Height          =   300
         Left            =   1605
         TabIndex        =   8
         Top             =   165
         Width           =   10050
      End
      Begin VB.Label LbDescCuenta 
         Caption         =   "Cuenta Contable "
         Height          =   180
         Left            =   225
         TabIndex        =   7
         Top             =   210
         Width           =   1365
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5460
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo1.frx":233A
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo1.frx":287E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo1.frx":2C10
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo1.frx":2D94
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo1.frx":31E8
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo1.frx":3300
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo1.frx":3844
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo1.frx":3D88
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo1.frx":3E9C
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo1.frx":3FB0
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo1.frx":4404
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo1.frx":4570
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo1.frx":4AB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo1.frx":4DD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo1.frx":5164
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHojaTrabajo1.frx":54F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar Asiento"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Object.ToolTipText     =   "Configurar Formatos"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Ver Detalle"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Exportar MSExcel"
      End
   End
End
Attribute VB_Name = "FrmHojaTrabajo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstTmp As New ADODB.Recordset
Dim SeEjecuto As Boolean

Dim mMesIni As Integer
Dim mMesFin As Integer
Dim BAND_INTERRUMPIR As Boolean '--interrumpir el procesos de la consulta

Dim mPosRegistro As Integer '--indica la posicion del numero de registro
Dim OrigFX As Long '--para mover el frame posicion horizontal
Dim OrigFY As Long '--para mover el frame posicion vertical
 

Sub Cargar(Indice As Integer)
    Dim Rst As New ADODB.Recordset
    Dim xFil As Integer
    Dim A As Integer
    
    PreparaRST_Tmp
    Fg1(Indice).Rows = 2
    DoEvents
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
            Fg1(Indice).TextMatrix(xFil, 7) = Format(Rst("debe") + NulosN(Fg1(Indice).TextMatrix(xFil, 3)), "0.00")
            Fg1(Indice).TextMatrix(xFil, 8) = Format(Rst("haber") + NulosN(Fg1(Indice).TextMatrix(xFil, 4)), "0.00")
            
            
            'saldo
            If NulosN(Fg1(Indice).TextMatrix(xFil, 7)) - NulosN(Fg1(Indice).TextMatrix(xFil, 8)) > 0 Then
                Fg1(Indice).TextMatrix(xFil, 9) = NulosN(Fg1(Indice).TextMatrix(xFil, 7)) - NulosN(Fg1(Indice).TextMatrix(xFil, 8))
                Fg1(Indice).TextMatrix(xFil, 9) = Format(Fg1(Indice).TextMatrix(xFil, 9), "0.00")
                Fg1(Indice).TextMatrix(xFil, 10) = "0.00"
            Else
                Fg1(Indice).TextMatrix(xFil, 9) = "0.00"
                Fg1(Indice).TextMatrix(xFil, 10) = NulosN(Fg1(Indice).TextMatrix(xFil, 8)) - NulosN(Fg1(Indice).TextMatrix(xFil, 7))
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
        xTotal1 = xTotal1 + NulosN(Fg1(Indice).TextMatrix(A, 3))
        xTotal2 = xTotal2 + NulosN(Fg1(Indice).TextMatrix(A, 4))
        xTotal3 = xTotal3 + NulosN(Fg1(Indice).TextMatrix(A, 5))
        xTotal4 = xTotal4 + NulosN(Fg1(Indice).TextMatrix(A, 6))
        xTotal5 = xTotal5 + NulosN(Fg1(Indice).TextMatrix(A, 7))
        xTotal6 = xTotal6 + NulosN(Fg1(Indice).TextMatrix(A, 8))
        xTotal7 = xTotal7 + NulosN(Fg1(Indice).TextMatrix(A, 9))
        xTotal8 = xTotal8 + NulosN(Fg1(Indice).TextMatrix(A, 10))
        xTotal9 = xTotal9 + NulosN(Fg1(Indice).TextMatrix(A, 11))
        xTotal10 = xTotal10 + NulosN(Fg1(Indice).TextMatrix(A, 12))
        xTotal11 = xTotal11 + NulosN(Fg1(Indice).TextMatrix(A, 13))
        xTotal12 = xTotal12 + NulosN(Fg1(Indice).TextMatrix(A, 14))
        xTotal13 = xTotal13 + NulosN(Fg1(Indice).TextMatrix(A, 15))
        xTotal14 = xTotal14 + NulosN(Fg1(Indice).TextMatrix(A, 16))
        xTotal15 = xTotal15 + NulosN(Fg1(Indice).TextMatrix(A, 17))
        xTotal16 = xTotal16 + NulosN(Fg1(Indice).TextMatrix(A, 18))
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
        
    If NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 10)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 9)) > 0 Then
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 9) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 9)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 10))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 9) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 9), "0.00")
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 10) = "0.00"
    Else
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 9) = "0.00"
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 10) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 10)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 9))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 10) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 10), "0.00")
    End If

    If NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 12)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 11)) > 0 Then
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 11) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 12)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 11))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 11) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 11), "0.00")
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 12) = "0.00"
    Else
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 11) = "0.00"
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 12) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 11)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 12))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 12) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 12), "0.00")
    End If
    
    If NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 14)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 13)) > 0 Then
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 13) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 14)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 13))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 13) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 13), "0.00")
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 14) = "0.00"
    Else
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 13) = "0.00"
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 14) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 13)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 14))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 14) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 14), "0.00")
    End If
    
    If NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 16)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 15)) > 0 Then
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 15) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 16)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 15))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 15) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 15), "0.00")
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 16) = "0.00"
    Else
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 15) = "0.00"
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 16) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 15)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 16))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 16) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 16), "0.00")
    End If
    
    If NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 18)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 17)) > 0 Then
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 17) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 18)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 17))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 17) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 17), "0.00")
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 18) = "0.00"
    Else
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 17) = "0.00"
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 18) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 17)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 18))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 18) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 18), "0.00")
    End If
    
    Fg1(Indice).Rows = Fg1(Indice).Rows + 1
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 2) = "S U M A S  I G U A L E S ==>"
    
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 9) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 9)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 9))
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 10) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 10)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 10))
    
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 11) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 11)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 11))
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 12) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 12)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 12))

    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 13) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 13)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 13))
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 14) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 14)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 14))

    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 15) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 15)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 15))
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 16) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 16)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 16))

    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 17) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 17)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 17))
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 18) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 18)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 18))
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
    mMesIni = SeleccionaMes(xCon)
    LblPerIni.Caption = Busca_Codigo(mMesIni, "id", "descripcion", "con_meses", "N", xCon)
End Sub

Private Sub cmd_periodo2_Click()
    mMesFin = SeleccionaMes(xCon)
    LblPerFin.Caption = Busca_Codigo(mMesFin, "id", "descripcion", "con_meses", "N", xCon)
End Sub


Private Sub Fg1_DblClick(Index As Integer)
    If Fg1(Index).Row < Fg1(Index).FixedRows Or Index <> 0 Then Exit Sub
    If Fg1(Index).Row >= Fg1(Index).Rows - 3 Then Exit Sub
    '--mostrando la ventana del detalle
    pHabilitarBotonEditor True, Index

End Sub

Private Sub Fg1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And Index = 0 Then
        PopupMenu menu1
    End If
End Sub

Private Sub fg2_DblClick()
    '--mostrar el asiento
    If Fg2.Rows <= Fg2.FixedRows Then Exit Sub
    Dim xfrm As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    xfrm.AsientoVer xCon, Fg2.TextMatrix(Fg2.Row, mPosRegistro)
    Set xfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()

    If SeEjecuto = False Then
    
        SeEjecuto = True
        TabOne1.CurrTab = 0
        Setea
        
        TxtIdMon.Text = 1
        TxtIdMon_Validate False
        
        opt_fecha(0).Value = True
        TxtLibro.Text = ""
        LblIdLibro.Caption = 0
        
        LimpiaText txt_cb
        LimpiaText lbl_cod
        LimpiaText lbl_cb
        
    End If
End Sub

Sub Setea()
    'usamos la columna 19 para almacenar el destino de cada cuenta en la hoja de trabajo
    Dim A As Integer
    Dim B As Integer
    
    '--buscar los registros
    Fg2.AutoSearch = flexSearchFromTop
    Fg1(0).AutoSearch = flexSearchFromTop
    Fg1(1).AutoSearch = flexSearchFromTop
    Fg1(2).AutoSearch = flexSearchFromTop
    
    For A = 0 To 3
'         Fg1(A).GridLines = flexGridNone
         Fg1(A).ColWidth(19) = 0
         Fg1(A).ColWidth(20) = 0
         Fg1(A).ColWidth(21) = 0
         Fg1(A).Rows = 2
         Fg1(A).RowHeight(0) = 350
         Fg1(A).RowHeight(1) = 300
         Fg1(A).TextMatrix(0, 1) = "          1"
         Fg1(A).TextMatrix(1, 1) = "          1"
         Fg1(A).TextMatrix(0, 1) = "Nº Cuenta"
         Fg1(A).TextMatrix(1, 1) = "Nº Cuenta"
         Fg1(A).TextMatrix(0, 2) = "Descripción"
         Fg1(A).TextMatrix(1, 2) = "Descripción"
         
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
             .Cell(flexcpText, 0, 11, 0, 12) = "Cuentas del Balance"
             .Cell(flexcpText, 0, 13, 0, 14) = "Transferencias y Canc."
             .Cell(flexcpText, 0, 15, 0, 16) = "Resultados x Naturaleza"
             .Cell(flexcpText, 0, 17, 0, 18) = "Resultados x Función"
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
         
         Fg1(A).TextMatrix(1, 19) = "IdDest"
         Fg1(A).TextMatrix(1, 20) = "IdDest2"
         Fg1(A).TextMatrix(1, 21) = "IdCta"
         
         For B = 3 To 18
             Fg1(A).ColAlignment(B) = flexAlignRightCenter
         Next B
        
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
                    .Cells(xFilas, B + 1) = NulosN(Fg1(Indice).TextMatrix(A, B))
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


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    ElseIf KeyCode = vbKeyF8 Then
        pConsultar
    End If
End Sub

Private Sub Form_Load()
    TxtFchIni.Valor = Date
    TxtFchFin.Valor = Date
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    
    Frame3.BackColor = &H8000000F
    Frame5.BackColor = &H8000000F
    Frame7.BackColor = &H8000000F
    
    SeEjecuto = False
    LblDescCuenta.Caption = ""
    
    LblPerIni.Caption = ""
    LblPerFin.Caption = ""
    
    TabOne1.CurrTab = 0
    
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
  
    If Me.Height > 3000 Then
        TabOne1.Top = 1500
        TabOne1.Width = Me.Width - 150
        TabOne1.Height = Me.Height - 1900
        '--detalle
        Fg1(1).Width = Me.Width - 200
        Fg1(1).Height = Me.Height - 3000
        '--analisis
        Fg1(3).Width = Me.Width - 200
        Fg1(3).Height = Me.Height - 3000
    End If
   
End Sub

Private Sub Menu1_1_Click()
    Fg1_DblClick TabOne1.CurrTab
End Sub

Private Sub Menu1_3_Click()
    pExportar TabOne1.CurrTab
End Sub

Private Sub opt_fecha_Click(Index As Integer)
    If Index = 0 Then
        Frame7.Visible = False
    Else
        Frame7.Top = 240
        Frame7.Visible = True
        
    End If
End Sub


Private Sub pic_Click(Index As Integer)
    pHabilitarBotonEditor False, 0
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    BAND_INTERRUMPIR = False
    
    If Button.Index = 1 Then pConsultar
    
    If Button.Index = 3 Then pBuscarAsiento
    
    'ExportarComprasExcel TabOne1.CurrTab
    If Button.Index = 5 Then pExportar TabOne1.CurrTab
        
    If Button.Index = 6 And TabOne1.CurrTab <> 3 Then pImprimir
    
    If Button.Index = 7 And TabOne1.CurrTab <> 3 Then Configurar

    If Button.Index = 9 Then Unload Me

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


'***********************************************************************************************
'***********************************************************************************************

Private Sub CmdBusProv_Click()
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
   
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "5500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, "SELECT * FROM mae_libros  where activo = -1 ORDER BY descripcion ", xCampos(), "Buscando Libro Contable", "descripcion", "descripcion", Principio
    If xRs.State = 1 Then
        TxtLibro.Text = NulosC(xRs("descripcion"))
        LblIdLibro.Caption = NulosC(xRs("id"))
    End If
    Set xRs = Nothing
End Sub


Private Sub TxtLibro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtLibro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusProv_Click
    End If
End Sub

Private Sub ChkLibro_Click()
    If ChkLibro.Value = 1 Then
        TxtLibro.Enabled = True
        CmdBusProv.Enabled = True
    Else
        TxtLibro.Enabled = False
        CmdBusProv.Enabled = False
        TxtLibro.Text = ""
        LblIdLibro.Caption = 0
    End If
End Sub

'***********************************************************************************************
'***********************************************************************************************


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
    ElseIf Button.Index = 3 Then '--exportar
        Dim xFun As New SGI2_funciones.formularios
'        xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1, lblCuenta.Caption, "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Expresado en " & LblMoneda.Caption, "Diario - Detalle"      ', Rst, ""
    
        Dim nPeriodo As String
        Dim nTitulo1 As String
        If opt_fecha(0).Value = True Then  '--por fecha
            If CDate(Me.TxtFchIni.Valor) <> CDate(Me.TxtFchFin.Valor) Then
                nPeriodo = " Del: " + CStr(TxtFchIni.Valor) + " Al: " + CStr(TxtFchFin.Valor)
            Else
                nPeriodo = "Al: " + CStr(TxtFchIni.Valor)
            End If
        Else '--por periodo
            If mMesIni = mMesFin Then
                nPeriodo = "Periodo:  " + LblPerIni.Caption
            Else
                nPeriodo = "Periodo:  De " + LblPerIni.Caption & " A " & LblPerFin.Caption
            End If
            
        End If
        nTitulo1 = "(Expresado en " & LblMoneda.Caption & ")"
    

        GRID_EXPORTAR_MSEXCELTMP Fg2, xCon, flexFileCustomText, True, lblCuenta.Caption, nPeriodo, nTitulo1
        
        Set xFun = Nothing

    ElseIf Button.Index = 4 Then '--imprimir
    ElseIf Button.Index = 6 Then
        pHabilitarBotonEditor False, 0
    End If
End Sub

Private Sub TxtIdMon_Change()
    If Trim(TxtIdMon.Text) = "" Then LblMoneda.Caption = ""
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtIdMon_Validate(Cancel As Boolean)
    If NulosC(TxtIdMon.Text) <> "" Then
        LblMoneda.Caption = Busca_Codigo(TxtIdMon.Text, "id", "descripcion", "mae_moneda", "N", xCon)
        If NulosC(LblMoneda.Caption) = "" Then
            TxtIdMon.Text = ""
        End If
    End If
End Sub

Private Sub CmdBusMon_Click()
    
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre":      xCampos(0, 1) = "descripcion":     xCampos(0, 2) = "4500":      xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":   xCampos(1, 1) = "id":              xCampos(1, 2) = "500":      xCampos(1, 3) = "N"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, "SELECT * FROM mae_moneda ORDER BY descripcion ;", xCampos(), "Buscando Moneda", "descripcion", "descripcion", Principio
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.RecordCount = 0 Then GoTo SALIR
    TxtIdMon.Text = xRs("id") & ""
    LblMoneda.Caption = xRs("descripcion") & ""
    
SALIR:
    Set xRs = Nothing
End Sub

Private Sub Procesar1(Indice As Integer)
    Dim A As Integer
    
    Label6.Caption = "Procesando Consulta"

    DoEvents
    '--colocar el saldo correspondiente al periodo de consulta en el encabezado
    If opt_fecha(0).Value = True Then
        Fg1(Indice).Cell(flexcpText, 0, 9, 0, 10) = "Saldos Al " & TxtFchFin.Valor
    Else
        Fg1(Indice).Cell(flexcpText, 0, 9, 0, 10) = "Saldos A " & LblPerFin.Caption
    End If
    
    '--limpiar informe
    Fg1(Indice).Rows = Fg1(Indice).FixedRows
    
    Cargar1 Indice
    
    DoEvents
    If BAND_INTERRUMPIR = True Then Exit Sub
    '--muestra totales
    Totalizar1 Indice
    
    DoEvents
    
    Fg1(Indice).FrozenCols = 2

    If Indice = 0 Then
        GRID_COLOR_FONDO Fg1(Indice), Fg1(Indice).FixedRows, 1, Fg1(Indice).Rows - 1, 2, &HCEFFFE  '&HECFFFF
        GRID_COLOR_FONDO Fg1(Indice), Fg1(Indice).FixedRows, 5, Fg1(Indice).Rows - 1, 6, &HCEFFFE
        GRID_COLOR_FONDO Fg1(Indice), Fg1(Indice).FixedRows, 9, Fg1(Indice).Rows - 1, 10, &HCEFFFE
        GRID_COLOR_FONDO Fg1(Indice), Fg1(Indice).FixedRows, 13, Fg1(Indice).Rows - 1, 14, &HCEFFFE
        GRID_COLOR_FONDO Fg1(Indice), Fg1(Indice).FixedRows, 17, Fg1(Indice).Rows - 1, 18, &HCEFFFE
        GRID_COLOR_FONDO Fg1(Indice), Fg1(Indice).Rows - 3, 1, Fg1(Indice).Rows - 1, Fg1(Indice).Cols - 1, &HC8FDD3 '&HE0FEE7
        
    Else
        GRID_COLOR_FONDO Fg1(Indice), Fg1(Indice).FixedRows, 1, Fg1(Indice).Rows - 1, 2, &HD5DDFD
        GRID_COLOR_FONDO Fg1(Indice), Fg1(Indice).FixedRows, 5, Fg1(Indice).Rows - 1, 6, &HD5DDFD
        GRID_COLOR_FONDO Fg1(Indice), Fg1(Indice).FixedRows, 9, Fg1(Indice).Rows - 1, 10, &HD5DDFD
        GRID_COLOR_FONDO Fg1(Indice), Fg1(Indice).FixedRows, 13, Fg1(Indice).Rows - 1, 14, &HD5DDFD
        GRID_COLOR_FONDO Fg1(Indice), Fg1(Indice).FixedRows, 17, Fg1(Indice).Rows - 1, 18, &HD5DDFD
        GRID_COLOR_FONDO Fg1(Indice), Fg1(Indice).Rows - 3, 1, Fg1(Indice).Rows - 1, Fg1(Indice).Cols - 1, &HFBE9C6 '&HE0FEE7
    
    End If
        
        
    Fg1(Indice).AutoSizeMode = flexAutoSizeColWidth
    
    '-totalizando
    For A = 3 To 18
        
        '--ajustando las columnas de acuerdo a los importes
        Fg1(Indice).AutoSize A
        If Fg1(Indice).ColWidth(A) < 850 Then Fg1(Indice).ColWidth(A) = 880
    Next A
    '-------------------------
        
End Sub






Private Sub Cargar1(Indice As Integer)
    '===================================================================================================
    'creado: 25/12/08
    'Propósito: Mostrar la información del balance de comprobacion
    '
    'Entradas:  Indice = Tipo de Reporte
    '
    'Resultados: balance de comprobacion libros en pantalla
    '
    'Modificado: 18/03/10 Por Johan Castro
    '           Mostras reporte a 2 ó 3 digitos
    '           09/04/10 Por Johan Castro
    '           Considerar en reporte asientos de cierre
    '===================================================================================================
    'LEYENDA:
    'SI: Saldos Iniciales
    'MP: Movimientos del Periodo
    'SM: Sumas del Mayor
    'SA: Saldos Al
    'CB: Cuentas de Balance
    'CT: Cuentas de Transferencia
    'GN: Ganancias por Naturaleza
    'GF: Ganancias por Funcion


    Dim nSQL As String
    Dim Rst As New ADODB.Recordset
    Dim nSQLIdLibro As String
    Dim nSQLTipoPersona As String
    Dim nSQLAjuste  As String '--sentencia sql para considera los registros del diario se ajuste por diferencia de cambio
    Dim nSQLCierre As String '--sentencia sql para no mostrar el cierre
    
    Dim FchIni, FchFin As Date
    
    Dim mDigitos As Integer '--cantidad de digitos que se muestra en HT
    
    '--identificar la cantidad de digitos a mostrar en informe
    If TabOne1.CurrTab = 1 Then
        If Opt2Digito.Value = True Then
            mDigitos = 2
        ElseIf Opt3Digito.Value = True Then
            mDigitos = 3
        Else
            '--por defecto
            mDigitos = 2
        End If
    End If
    
    
    Frame11.Left = 3120
    Frame11.Top = 3930
    Frame11.Visible = True
    ProgressBar1.Min = 0
    ProgressBar1.Value = 0

    
    If NulosN(LblIdLibro.Caption) <> 0 Then nSQLIdLibro = " and con_diario.idlib=" & LblIdLibro
    '--si selecciona un tipo de persona
    If lbl_cod(1).Caption <> "" Then
        If NulosN(lbl_cod(0).Caption) <> 0 Then
            nSQLTipoPersona = " and con_diario.ridtipper = " & lbl_cod(0).Caption & " and con_diario.ridper = " & lbl_cod(1).Caption
        Else
            '--buscar
            nSQLTipoPersona = " and " & Replace(lbl_cod(1).Caption, "cod", "con_diario.ridper")
        End If
    End If
    
        '--para ajuste por diferencia de cambio
    nSQLAjuste = " AND (con_diario.ajuste in (0, " & NulosN(TxtIdMon.Text) & ") ) "
    '-----------------------------------------------
    nSQLCierre = " AND (con_diario.idmes<>13) "
    '-----------------------------------------------
    
    
    DoEvents
    
    '--19/04/09
    '--se cambia los saldos iniciales solo debera de mostrar debe o haber
    '-- IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol) AS SIDebSol,
    '-- IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol) AS SIHabSol,
    nSQL = "SELECT con_planctas.cuenta, con_planctas.descripcion, " _
        + vbCr + " IIf(((IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))-(IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol)))>0,((IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))-(IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))),0) AS SIDebSol, " _
        + vbCr + " IIf(((IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))-(IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol)))>0,((IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))-(IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))),0) AS SIHabSol, " _
        + vbCr + " IIf(MovPeriodo.DebSol Is Null,0,MovPeriodo.DebSol) AS MPDebSol, " _
        + vbCr + " IIf(MovPeriodo.HabSol Is Null,0,MovPeriodo.HabSol) AS MPHabSol, " _
        + vbCr + " [SIDebSol]+[MPDebSol] AS SMDebSol, " _
        + vbCr + " [SIHabSol]+[MPHabSol] AS SMHabSol, " _
        + vbCr + " IIf((SMDebSol-SMHabSol)>0,(SMDebSol-SMHabSol),0) AS SADebSol, " _
        + vbCr + " IIf((SMHabSol-SMDebSol)>0,(SMHabSol-SMDebSol),0) AS SAHabSol, " _
        + vbCr + " IIf(con_planctas.iddes=1 Or con_planctas.iddes2=1,SADebSol,0) AS CBDebSol, " _
        + vbCr + " IIf(con_planctas.iddes=1 Or con_planctas.iddes2=1,SAHabSol,0) AS CBHabSol, " _
        + vbCr + " IIf(con_planctas.iddes=4 Or con_planctas.iddes2=4,SADebSol,0) AS CTDebSol, " _
        + vbCr + " IIf(con_planctas.iddes=4 Or con_planctas.iddes2=4,SAHabSol,0) AS CTHabSol, " _
        + vbCr + " IIf(con_planctas.iddes=2 Or con_planctas.iddes2=2,SADebSol,0) AS GNDebSol, " _
        + vbCr + " IIf(con_planctas.iddes=2 Or con_planctas.iddes2=2,SAHabSol,0) AS GNHabSol, " _
        + vbCr + " IIf(con_planctas.iddes=3 Or con_planctas.iddes2=3,SADebSol,0) AS GFDebSol, " _
        + vbCr + " IIf(con_planctas.iddes=3 Or con_planctas.iddes2=3,SAHabSol,0) As GFHabSol, " _
        + vbCr + " con_planctas.iddes,con_planctas.iddes2,con_planctas.id AS IdCta "
    
    '--agregar a reporte columna para identificar los numeros de digitos
    If TabOne1.CurrTab = 1 Then
        nSQL = nSQL & ", Left([con_planctas].[cuenta]," & mDigitos & ") as grupo"
    End If
    
    
    If NulosN(TxtIdMon.Text) = 2 Then
        nSQL = Replace(nSQL, "DebSol", "DebDol")
        nSQL = Replace(nSQL, "HabSol", "HabDol")
    End If
    '--movimientos del periodo
    nSQL = nSQL _
        + vbCr + " FROM (con_planctas LEFT JOIN " _
        + vbCr + " ( " _
        + vbCr + " SELECT con_planctas.id AS IdCta, con_planctas.cuenta, con_planctas.descripcion, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS DebSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS HabSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS DebDol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabdol)) As HabDol " _
        + vbCr + " FROM (con_planctas RIGHT JOIN con_diario ON con_planctas.id=con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc=con_tc.fecha " _
        + vbCr + " WHERE (((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) " & nSQLIdLibro & nSQLTipoPersona & nSQLAjuste _
        + vbCr + " GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion " _
        + vbCr + " ORDER BY con_planctas.cuenta " _
        + vbCr + " ) AS MovPeriodo ON con_planctas.id = MovPeriodo.IdCta) " _
        + vbCr + " Left Join "
    
    '--saldos iniciales
    nSQL = nSQL _
        + vbCr + " ( " _
        + vbCr + " SELECT con_planctas.id AS IdCta, con_planctas.cuenta, con_planctas.descripcion, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS DebSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS HabSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS DebDol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabdol)) As HabDol " _
        + vbCr + " FROM (con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
        + vbCr + " WHERE  ((con_diario.fchasi) Is Null Or (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "')) " & nSQLIdLibro & nSQLTipoPersona & nSQLAjuste & nSQLCierre _
        + vbCr + " GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion " _
        + vbCr + " ORDER BY con_planctas.cuenta " _
        + vbCr + " ) AS SaldosIni "
    
    '--filtro para el where
    nSQLIdLibro = nSQLIdLibro & nSQLTipoPersona & nSQLAjuste & " AND (  (((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')))  OR  (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "')  OR  (con_diario.fchasi) is null  )"


    nSQL = nSQL _
        + vbCr + " ON con_planctas.id = SaldosIni.IdCta " _
        + vbCr + " WHERE con_planctas.id In (SELECT con_diario.idcue FROM con_diario " & IIf(nSQLIdLibro <> "", "WHERE " & Mid(nSQLIdLibro, 5), "") & "   ) " _
        + vbCr + " ORDER BY con_planctas.cuenta; "

    '--si seleccionar por periodo
    If opt_fecha(1).Value = True Then
        '--movimiento del periodo
        nSQL = Replace(nSQL, "(((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')))", "( con_diario.idmes>=" & mMesIni & " And con_diario.idmes <= " & mMesFin & " )")
        '--saldos iniciales
        nSQL = Replace(nSQL, "((con_diario.fchasi) Is Null Or (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "'))", "con_diario.idmes < " & mMesIni)
        '--filtro para where
        nSQL = Replace(nSQL, "OR  (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "')  OR  (con_diario.fchasi) is null ", " OR con_diario.idmes <= " & mMesIni)
        
    End If
    
    
    
    Me.MousePointer = vbHourglass
    DoEvents
    RST_Busq Rst, nSQL, xCon
    
    
    Dim RstGrp As New ADODB.Recordset
    Dim mCol&

    '--verificar si se muestra en grupos
    If TabOne1.CurrTab = 1 Then
        '--NOTA: QUEDA PENDIENTE QUE MUESTRE LAS CUENTAS QUE NO ESTAN EL EL GRUPO QUE SE SOLICITA
        '--EJ. CONSULTA A 3 DIGITOS: PERO HAY 3 CUENTAS QUE SOLO LLEGAN A 2 DIGITOS; EN ESTO REPORTE NO SE MUESTRA
        '--FALTA CORREGIR ESTE DETALLE
    
        '--obtener listado de cuentas en grupo
        nSQL = "SELECT con_planctas.cuenta, con_planctas.descripcion " _
            + vbCr + " From con_planctas " _
            + vbCr + " Where Len([con_planctas].[cuenta]) = " & mDigitos & " " _
            + vbCr + " ORDER BY con_planctas.cuenta; "
        '--cargar consulta
        RST_Busq RstGrp, nSQL, xCon
        
        '--indicar cantidad de registros a mostrar en barra de progreso
        If RstGrp.RecordCount <> 0 Then ProgressBar1.Max = RstGrp.RecordCount
        
        Do While Not RstGrp.EOF
            
            '-----------------------------------------------
            ProgressBar1.Value = ProgressBar1.Value + 1
            
            '--aplicar filtro segun grupo
            Rst.Filter = "grupo='" & NulosC(RstGrp("cuenta")) & "'"
            '--verificar si hay registros en grupo
            If Rst.RecordCount <> 0 Then
                '--proceder a cargar los datos
                    DoEvents
                    Fg1(Indice).Rows = Fg1(Indice).Rows + 1
                    
                    DoEvents
                    '--SI SE NTERRUMPE EL PROCESO => SALIR
                    If BAND_INTERRUMPIR = True Then GoTo SALIR
                    DoEvents
                    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 1) = NulosC(RstGrp("cuenta"))
                    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 2) = NulosC(RstGrp("descripcion"))
                    
                    For mCol = 2 To Rst.Fields.Count - 1
                        '--posicionar en la primera fila
                        Rst.MoveFirst
                        
                        If mCol <= 1 Then
                            
                        ElseIf mCol < 18 Then
                            '--importes
                            If Chk0Decimales.Value = 1 Then
                                Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, mCol + 1) = Format(RstRegistroSumar(Rst, Rst(mCol).Name), "###,###,##0")
                            Else
                                Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, mCol + 1) = Format(RstRegistroSumar(Rst, Rst(mCol).Name), FORMAT_MONTO)
                            End If
                            
                        Else
                                                   
                        
                        End If
                    Next mCol
            End If
            
            '--ir al siguiente registro
            RstGrp.MoveNext
        Loop
        
        
    Else
        If Rst.RecordCount <> 0 Then ProgressBar1.Max = Rst.RecordCount
        '--proceder a cargar los datos
        Do While Not Rst.EOF
            DoEvents
            Fg1(Indice).Rows = Fg1(Indice).Rows + 1
            '-----------------------------------------------
            ProgressBar1.Value = ProgressBar1.Value + 1
            DoEvents
            '--SI SE NTERRUMPE EL PROCESO => SALIR
            If BAND_INTERRUMPIR = True Then GoTo SALIR
            DoEvents
            For mCol = 0 To Rst.Fields.Count - 1
                
                If mCol <= 1 Then
                    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, mCol + 1) = NulosC(Rst(mCol))
                ElseIf mCol < 18 Then
                    '--importes
                    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, mCol + 1) = Format(Rst(mCol), FORMAT_MONTO)
                Else
                    '--iddest1,iddest2,idcta
                    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, mCol + 1) = NulosN(Rst(mCol))
                
                
                End If
            Next mCol
            Rst.MoveNext
        Loop
    End If
SALIR:
    Frame11.Visible = False
    Me.MousePointer = vbDefault
    
End Sub




Private Sub Totalizar1(Indice As Integer)
    Dim A As Integer
    
    
    Fg1(Indice).Rows = Fg1(Indice).Rows + 1
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 2) = "T O T A L E S ==>"
    
    
    '-totalizando
    For A = 3 To 18
        
        
        If Indice = 1 And Chk0Decimales = 1 Then
            Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, A) = Format(GRID_SUMAR_COL(Fg1(Indice), A), "###,###,##0")
        Else
            Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, A) = Format(GRID_SUMAR_COL(Fg1(Indice), A), FORMAT_MONTO)
        End If
        
    Next A
    '-------------------------
    
    
    Fg1(Indice).Rows = Fg1(Indice).Rows + 1
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 2) = "R E S U L T A D O ==>"
        
    If NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 10)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 9)) > 0 Then
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 9) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 9)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 10))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 9) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 9), FORMAT_MONTO)
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 10) = "0.00"
    Else
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 9) = "0.00"
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 10) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 10)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 9))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 10) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 10), FORMAT_MONTO)
    End If

    If NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 12)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 11)) > 0 Then
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 11) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 12)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 11))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 11) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 11), FORMAT_MONTO)
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 12) = "0.00"
    Else
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 11) = "0.00"
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 12) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 11)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 12))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 12) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 12), FORMAT_MONTO)
    End If
    
    If NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 14)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 13)) > 0 Then
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 13) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 14)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 13))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 13) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 13), FORMAT_MONTO)
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 14) = "0.00"
    Else
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 13) = "0.00"
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 14) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 13)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 14))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 14) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 14), FORMAT_MONTO)
    End If
    
    If NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 16)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 15)) > 0 Then
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 15) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 16)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 15))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 15) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 15), FORMAT_MONTO)
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 16) = "0.00"
    Else
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 15) = "0.00"
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 16) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 15)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 16))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 16) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 16), FORMAT_MONTO)
    End If
    
    If NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 18)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 17)) > 0 Then
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 17) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 18)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 17))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 17) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 17), FORMAT_MONTO)
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 18) = "0.00"
    Else
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 17) = "0.00"
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 18) = NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 17)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 18))
        Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 18) = Format(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 18), FORMAT_MONTO)
    End If
    
    Fg1(Indice).Rows = Fg1(Indice).Rows + 1
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 2) = "S U M A S  I G U A L E S ==>"
    
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 9) = Format(NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 9)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 9)), FORMAT_MONTO)
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 10) = Format(NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 10)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 10)), FORMAT_MONTO)
    
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 11) = Format(NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 11)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 11)), FORMAT_MONTO)
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 12) = Format(NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 12)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 12)), FORMAT_MONTO)

    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 13) = Format(NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 13)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 13)), FORMAT_MONTO)
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 14) = Format(NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 14)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 14)), FORMAT_MONTO)

    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 15) = Format(NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 15)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 15)), FORMAT_MONTO)
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 16) = Format(NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 16)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 16)), FORMAT_MONTO)

    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 17) = Format(NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 17)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 17)), FORMAT_MONTO)
    Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, 18) = Format(NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 3, 18)) + NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 2, 18)), FORMAT_MONTO)
    
End Sub


Private Sub pExportar(Indice As Integer)
    Dim xFun As New SGI2_funciones.formularios
    Dim Rst As New ADODB.Recordset
    
    xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1(Indice), "HOJA DE TRABAJO", "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor, "Expresado en : " & LblMoneda.Caption, "Balance de Comprobación"
    Set xFun = Nothing
    
End Sub



'*******************************************************************************************

Private Sub cb_Click(Index As Integer)
    Dim xCampos() As String
    Dim nCampoBusca As String
    Dim nSQL As String
    Dim nTitulo As String
    On Error GoTo error
    Select Case Index
        Case 0 '--tipo personal
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":           xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
            nTitulo = "Buscando Tipo Persona"
            nSQL = "SELECT tes_tipopers.id, tes_tipopers.descripcion as nombre, tes_tipopers.id AS cod FROM tes_tipopers; "
            txt_cb(1).Text = ""
        Case 1 '--tipo persona
            Select Case NulosN(lbl_cod(0).Caption)
                Case 0: '--ninguno
                    MsgBox "Seleccione el Tipo", vbExclamation, xTitulo
                    txt_cb(0).SetFocus
                    
                    
                    Exit Sub
                Case 1 '--proveedor
                    nSQL = "SELECT mae_prov.numruc AS numdoc, mae_prov.nombre, mae_prov.id AS cod  FROM mae_prov  Where ((mae_prov.activo) = -1) ; "

                Case 2 '--cliente'
                    nSQL = "SELECT mae_cliente.numruc AS numdoc, mae_cliente.nombre, mae_cliente.id AS cod   FROM mae_cliente   Where ((mae_cliente.activo) = -1) ;"

                Case 3 '--empleado
                    nSQL = "  SELECT pla_empleados.numdoc AS numdoc, pla_empleados.nombre, pla_empleados.id AS cod  FROM pla_empleados    Where   pla_empleados.numdoc is not null and pla_empleados.numdoc<>''  ;"


                Case 4 '--otros
                    
                    Exit Sub
                Case 5 '--bancos
                    nSQL = "SELECT mae_bancos.numruc AS numdoc, mae_bancos.descripcion as nombre, mae_bancos.id AS cod   FROM mae_bancos "
                    
                Case Else
                
                    Exit Sub
            End Select
            ReDim xCampos(3, 3) As String
            xCampos(0, 0) = "Nº Ruc":       xCampos(0, 1) = "numdoc":       xCampos(0, 2) = "1500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Nombres":      xCampos(1, 1) = "nombre":       xCampos(1, 2) = "5000":   xCampos(1, 3) = "C"
            xCampos(2, 0) = "Id":           xCampos(2, 1) = "cod":           xCampos(2, 2) = "700":    xCampos(2, 3) = "N"
            nTitulo = "Buscando " & lbl_cb(0).Caption

    End Select
    
    
    Dim RstTmp As New ADODB.Recordset
    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio

    If RstTmp.State = 0 Then GoTo SALIR
    If RstTmp.EOF = True Or RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index).Text = NulosC(RstTmp.Fields(0))  '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = NulosC(RstTmp.Fields(1)) '--NOMBRE
    lbl_cod(Index).Caption = NulosN(RstTmp.Fields(2)) '--CODIGO
    lbl_cb(Index).ToolTipText = NulosC(RstTmp.Fields(1))  '--NOMBRE
      

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
        If Index = 0 Then txt_cb(1).Text = ""
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
            SendKeys vbTab
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
        Case 0 '--tipo personal
            nSQL = "SELECT tes_tipopers.id, tes_tipopers.descripcion, tes_tipopers.id AS cod FROM tes_tipopers where tes_tipopers.id=" & NulosN(txt_cb(Index).Text)
        Case 1 '--tipo persona
            Select Case NulosN(lbl_cod(0).Caption)
                Case 0: '--ninguno
                    nSQL = "SELECT mae_prov.numruc AS numdoc, mae_prov.nombre, mae_prov.id AS cod FROM mae_prov  Where mae_prov.numruc='" & NulosN(txt_cb(Index).Text) & "' "
                    RST_Busq RstTmp, nSQL, xCon
                    If RstTmp.RecordCount <> 0 Then
                        txt_cb(Index).Text = NulosC(RstTmp.Fields(0))   '--TEXTO A MOSTRAR
                        lbl_cb(Index).Caption = NulosC(RstTmp.Fields(1)) '--NOMBRE
                        lbl_cod(Index).Caption = "( con_diario.ridtipper=1 and " & RstRegistroGenerarId(RstTmp, "cod", "", "in", True) & ")"
                    End If
                    
                    nSQL = " SELECT mae_cliente.numruc AS numdoc, mae_cliente.nombre, mae_cliente.id AS cod FROM mae_cliente Where mae_cliente.numruc ='" & NulosN(txt_cb(Index).Text) & "' "
                    RST_Busq RstTmp, nSQL, xCon
                    If RstTmp.RecordCount <> 0 Then
                        txt_cb(Index).Text = NulosC(RstTmp.Fields(0))   '--TEXTO A MOSTRAR
                        lbl_cb(Index).Caption = NulosC(RstTmp.Fields(1)) '--NOMBRE
                        If lbl_cod(Index).Caption <> "" Then lbl_cod(Index).Caption = lbl_cod(Index).Caption & " OR "
                        lbl_cod(Index).Caption = lbl_cod(Index).Caption & "( con_diario.ridtipper=2 and " & RstRegistroGenerarId(RstTmp, "cod", "", "in", True) & ")"
                    End If
                    
                    nSQL = " SELECT pla_empleados.numdoc AS numdoc, pla_empleados.nombre, pla_empleados.id AS cod FROM pla_empleados where  pla_empleados.numdoc='" & NulosN(txt_cb(Index).Text) & "'"
                    RST_Busq RstTmp, nSQL, xCon
                    If RstTmp.RecordCount <> 0 Then
                        txt_cb(Index).Text = NulosC(RstTmp.Fields(0))   '--TEXTO A MOSTRAR
                        lbl_cb(Index).Caption = NulosC(RstTmp.Fields(1)) '--NOMBRE
                        If lbl_cod(Index).Caption <> "" Then lbl_cod(Index).Caption = lbl_cod(Index).Caption & " OR "
                        lbl_cod(Index).Caption = lbl_cod(Index).Caption & "( con_diario.ridtipper=3 and " & RstRegistroGenerarId(RstTmp, "cod", "", "in", True) & ")"
                    End If
                    '--banco
                    nSQL = " SELECT mae_bancos.numruc AS numdoc, mae_bancos.descripcion as nombre, mae_bancos.id AS cod FROM mae_bancos Where mae_bancos.numruc ='" & NulosN(txt_cb(Index).Text) & "' "
                    RST_Busq RstTmp, nSQL, xCon
                    If RstTmp.RecordCount <> 0 Then
                        txt_cb(Index).Text = NulosC(RstTmp.Fields(0))   '--TEXTO A MOSTRAR
                        lbl_cb(Index).Caption = NulosC(RstTmp.Fields(1)) '--NOMBRE
                        If lbl_cod(Index).Caption <> "" Then lbl_cod(Index).Caption = lbl_cod(Index).Caption & " OR "
                        lbl_cod(Index).Caption = lbl_cod(Index).Caption & "( con_diario.ridtipper=5 and " & RstRegistroGenerarId(RstTmp, "cod", "", "in", True) & ")"
                    End If
                    
                    Set RstTmp = Nothing
                    Exit Sub
                Case 1 '--proveedor
                    nSQL = "SELECT mae_prov.numruc as numdoc, mae_prov.nombre, mae_prov.id as cod  FROM mae_prov WHERE (((mae_prov.activo)=-1)) and mae_prov.numruc ='" & NulosN(txt_cb(Index).Text) & "' "
                Case 2 '--cliente'
                    nSQL = " SELECT mae_cliente.numruc as numdoc, mae_cliente.nombre, mae_cliente.id AS cod From mae_cliente WHERE (((mae_cliente.activo)=-1)) and mae_cliente.numruc ='" & NulosN(txt_cb(Index).Text) & "' "
                Case 3 '--empleado
                    nSQL = "SELECT pla_empleados.numdoc, pla_empleados.nombre, pla_empleados.id AS cod FROM pla_empleados;"

                Case 4 '--otros
                    Exit Sub
                Case 5 '--bancos
                    nSQL = "SELECT mae_bancos.numruc as numdoc, mae_bancos.descripcion as nombre, mae_bancos.id as cod  FROM mae_bancos WHERE mae_bancos.numruc ='" & NulosN(txt_cb(Index).Text) & "' "
                    
                Case Else
                
                    Exit Sub
            End Select
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
        If NulosN(lbl_cod(0).Caption) = 0 Then
            lbl_cod(Index).Caption = RstRegistroGenerarId(RstTmp, "cod", "", "in", True)
        End If
    
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



'******************************************************************************

Private Sub pHabilitarBotonEditor(band As Boolean, Indice As Integer)
    '==================================================
    'Propósito: Mostrar el detalle de la cuenta
    '
    'Entradas:  band= puede ser true o false
    '
    'Resultados: Informacion detallada de la cuenta seleccionada
    '
    'Modificado: 30/05/11 Johan Castro
    '            Comentar lineas que bloquean objetos cuando se muestra el detalla
    '==================================================

    '--true muestra el ingreso de datos
    BAND_INTERRUMPIR = False
    
    FraDetalle.Visible = band
    If band = True Then
        TabOne1.CurrTab = 0
        FraDetalle.Top = 1470
        FraDetalle.Left = 60
        SetearCuadricula Fg2, 5, xCon, 1, 0, False
        lblCuenta.Caption = "DETALLE DE LA CUENTA  " & Fg1(Indice).TextMatrix(Fg1(Indice).Row, 1) & "  " & Fg1(Indice).TextMatrix(Fg1(Indice).Row, 2)
        DoEvents
    End If
    
    
'''    Toolbar1.Enabled = Not band
'''    TabOne1.Enabled = Not band
    
    If band = True Then pCargarDatosCuenta Indice
    
    
End Sub

Private Sub pCargarDatosCuenta(Indice As Integer)
    '===================================================================================================
    'Creado:     11/01/09
    'Propósito:  Mostrar información de la cuenta seleccionada
    'Indice:     index del array grilla
    'Resultados: Informacion detallada de la cuenta seleccionada
    '===================================================================================================
    Dim RstTmp2 As New ADODB.Recordset
    
    Dim nSQL As String
    Dim nSQLSaldo As String
    Dim nSQLWhere As String
    Dim nSQLCampos As String
    
    Dim mColDebe As Integer '--posicion de la columna debe
    Dim mColHaber As Integer '--posicion de la columna haber
    Dim mColSaldo As Integer '--posicion de la columna  saldo

    Dim nSQLAjuste As String

    Dim mColCampo As Long
    
    Dim mCol& '--indica la posicion del campo
    Dim xSaldo As Double
    Dim xTotal1, xTotal2 As Double
    Dim nTipoSaldo As String
    
    Frame11.Left = 3120
    Frame11.Top = 3930
    Frame11.Visible = True
    ProgressBar1.Min = 0
    ProgressBar1.Value = 0

    
    Me.MousePointer = vbHourglass
    Fg2.Rows = Fg2.FixedRows
    '---
    DoEvents
    '----------------------------------------------------------------------------------
    If NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Row, 21)) = 0 Then
        MsgBox "Seleccione la cuenta correctamente", vbExclamation, xTitulo
        Exit Sub
    End If
    '--colocando el saldo
    nSQL = "select * from con_planctas where con_planctas.id=" & Fg1(Indice).TextMatrix(Fg1(Indice).Row, 21)
    RST_Busq RstTmp2, nSQL, xCon
    
    If UCase(RstTmp2.Fields("tipsal") & "") = "D" Or NulosC(RstTmp2.Fields("tipsal")) = "" Then
        xSaldo = (NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Row, 3)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Row, 4)))
    Else
        xSaldo = (NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Row, 4)) - NulosN(Fg1(Indice).TextMatrix(Fg1(Indice).Row, 3)))
    End If
    nTipoSaldo = UCase(NulosC(RstTmp2.Fields("tipsal")))
    
    Set RstTmp2 = Nothing
    
    '----------------------------------------------------------------------------------
    
    '**********************************************************************************************
    nSQLCampos = fSetearCuadriculaColumna(xCon, 5)
    If nSQLCampos = "" Then Exit Sub
    nSQLCampos = "idcuenta,tipsal," & nSQLCampos

    '**********************************************************************************************
    
    If NulosN(LblIdLibro.Caption) <> 0 Then nSQLWhere = " and con_diario.idlib=" & LblIdLibro
    '--si selecciona un tipo de persona
    If lbl_cod(1).Caption <> "" Then
        If NulosN(lbl_cod(0).Caption) <> 0 Then
            nSQLWhere = nSQLWhere & " and con_diario.ridtipper = " & lbl_cod(0).Caption & " and con_diario.ridper = " & lbl_cod(1).Caption
        Else
            '--buscar
            nSQLWhere = nSQLWhere & " and " & Replace(lbl_cod(1).Caption, "cod", "con_diario.ridper")
        End If
    End If
    '**********************************************************************************************
    '--para ajuste por diferencia de cambio
    nSQLAjuste = " AND (con_diario.ajuste in (0, " & NulosN(TxtIdMon.Text) & ") ) "
    nSQLWhere = nSQLWhere & nSQLAjuste
    '-----------------------------------------------

    
    
    '--generando la consulta
    '--entes de 07/02/09

   '--08/03/09
   '--tomar tipo de cambio del diario cuando idlib = bancos y diversos
   nSQL = "SELECT con_diario.idcue AS idcuenta,con_planctas.tipsal,Format(con_diario.idmes,'00') & IIf(mae_libros.codsun Is Null,'',mae_libros.codsun) & con_diario.numasi AS registro, Format(con_diario.idmes,'00') AS mes, mae_libros.codsun AS libsun, CDbl(con_diario.numasi) AS corr, mae_libros.descripcion AS libdesc, con_diario.fchdoc AS fchope, con_diario.rglosaope as glosaope, con_diario.rglosa AS glosaref, con_diario.rregistro AS registroref, iif(con_diario.ridtipper=0,tes_documentos.abrev,mae_documento.abrev) AS tdocdesc, con_diario.rfchope AS fchdoc, con_diario.rnumerodoc AS numdoc, " _
            + vbCr + " IIf(con_diario.ridtipper=1,mae_prov.numruc,IIf(con_diario.ridtipper=2,mae_cliente.numruc,IIf(con_diario.ridtipper=3,pla_empleados.numdoc,IIf(con_diario.ridtipper=5,mae_bancos.numruc,'')))) AS numruc, " _
            + vbCr + " IIf(con_diario.ridtipper=1,mae_prov.nombre,IIf(con_diario.ridtipper=2,mae_cliente.nombre,IIf(con_diario.ridtipper=3,pla_empleados.apepat & ' ' & pla_empleados.apemat & ', ' & pla_empleados.nom,IIf(con_diario.ridtipper=5,mae_bancos.descripcion,'')))) AS apenom , mae_documento.codsun AS tdocsun, iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven)) AS tipcam, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo AS monope, mae_moneda_1.simbolo AS monref, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.impdebdol*tipcam),con_diario.impdebsol) AS impdebesol, " _
            + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.imphabdol*tipcam),con_diario.imphabsol) AS imphabersol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(tipcam=0 Or con_diario.impdebsol=0,0,(con_diario.impdebsol/tipcam))) AS impdebedol, " _
            + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(tipcam=0 Or con_diario.imphabsol=0,0,(con_diario.imphabsol/tipcam))) AS imphaberdol, " _
            + vbCr + " IIf(con_planctas.tipsal='D' or con_planctas.tipsal='',impdebsol-imphabsol,imphabsol-impdebsol) as impsalsol, " _
            + vbCr + " IIf(con_planctas.tipsal='D' or con_planctas.tipsal='',impdebdol-imphabdol,imphabdol-impdebdol) as impsaldol, " _
            + vbCr + " iif(con_diario.rnumerodoc1 is null,'',mae_documento_1.abrev) AS tdocdesc1, con_diario.rnumerodoc1 AS numdoc1, " _
            + vbCr + " tes_documentos_1.abrev AS tdocdesc2, con_diario.rfchope2 AS fchdoc2, con_diario.rnumerodoc2 AS numdoc2,con_diario.ridtipper2, iif(con_diario.ridtipper2<>5,'', mae_bancos_1.numruc ) AS numruc2,iif(con_diario.ridtipper2<>5,'',mae_bancos_1.descripcion ) AS apenom2 " _
            + vbCr + " FROM ((((((pla_empleados RIGHT JOIN (mae_cliente RIGHT JOIN ((((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN mae_documento ON con_diario.rtipdoc = mae_documento.id) LEFT JOIN mae_prov ON con_diario.ridper = mae_prov.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) ON mae_cliente.id = con_diario.ridper) ON pla_empleados.id = con_diario.ridper) LEFT JOIN tes_documentos ON con_diario.rtipdoc = tes_documentos.id) LEFT JOIN mae_bancos ON con_diario.ridper = mae_bancos.id) LEFT JOIN mae_bancos AS mae_bancos_1 ON con_diario.ridper2 = mae_bancos_1.id) LEFT JOIN mae_documento AS mae_documento_1 ON con_diario.rtipdoc1 = mae_documento_1.id) LEFT JOIN tes_documentos AS tes_documentos_1 ON con_diario.rtipdoc2 = tes_documentos_1.id) LEFT JOIN mae_moneda AS mae_moneda_1 ON con_diario.ridmon = mae_moneda_1.id "
   
   
    If opt_fecha(0).Value = True Then
        nSQL = nSQL + vbCr + " WHERE (con_diario.fchasi >=CDate('" + TxtFchIni.Valor + "') And con_diario.fchasi<=CDate('" + TxtFchFin.Valor + "')) and ( year(con_diario.fchasi)= " & AnoTra & " ) "
    Else
        nSQL = nSQL + vbCr + " WHERE ( con_diario.idmes >= " & mMesIni & " and con_diario.idmes <= " & mMesFin & " ) and con_diario.año = " & AnoTra & " "
    End If
    '--buscando solo la cuenta seleccionada
    nSQL = nSQL & " AND con_diario.idcue = " & Fg1(Indice).TextMatrix(Fg1(Indice).Row, 21) & nSQLWhere
    '---------------------------------------------------------------------
    nSQL = nSQL + vbCr + " ORDER BY con_planctas.cuenta ASC "

     '**********************************************************************************************
    '--remplazando segun la moneda seleccionada
    'If NulosN(TxtIdMon.Text) = 1 Then
    If NulosN(TxtIdMon.Text) = 1 Then
        nSQL = Replace(nSQL, "impdebesol", "debe")
        nSQL = Replace(nSQL, "imphabersol", "haber")
        nSQL = Replace(nSQL, "impsalsol", "saldo")
    Else
        nSQL = Replace(nSQL, "impdebedol", "debe")
        nSQL = Replace(nSQL, "imphaberdol", "haber")
        nSQL = Replace(nSQL, "impsaldol", "saldo")
    End If
    
    nSQL = "Select " & nSQLCampos & _
            vbCr + " from ( " _
            + vbCr + nSQL _
            + vbCr + ") as diario ORDER BY registro, ctanum ,numdoc"
    
    RST_Busq RstTmp2, nSQL, xCon
    '---------------------------------------------------------------------------
    '--obtener la posicione de las columnas debe,haber,saldo
    mCol = 0
    For mColCampo = 2 To RstTmp2.Fields.Count - 1
        mCol = mCol + 1
        Select Case LCase(RstTmp2.Fields(mColCampo).Name)
            Case "debe": mColDebe = mCol
            Case "haber": mColHaber = mCol
            Case "saldo": mColSaldo = mCol
            Case "registro": mPosRegistro = mCol
        End Select
    Next mColCampo


    Fg2.Rows = Fg2.Rows + 1
    Fg2.TextMatrix(Fg2.Rows - 1, 2) = "SALDO INICIAL"
    '----------------------------------------------------------------------------------
    If nTipoSaldo = "D" Or nTipoSaldo = "" Then
        Fg2.TextMatrix(Fg2.Rows - 1, mColDebe) = Format(xSaldo, FORMAT_MONTO)
    Else
        Fg2.TextMatrix(Fg2.Rows - 1, mColHaber) = Format(xSaldo, FORMAT_MONTO)
    End If
    FORMATO_CELDA Fg2, Fg2.Rows - 1, mColSaldo, , True, , Format(xSaldo, FORMAT_MONTO)
    

    
    '----------------------------------------------------------------------------------
    If RstTmp2.State = 0 Then GoTo SALIR
    If RstTmp2.BOF = True Or RstTmp2.EOF = True Or RstTmp2.RecordCount = 0 Then GoTo SALIR
    RstTmp2.MoveFirst
    '---------------------------------------------------------------------------
    
    DoEvents
    If RstTmp2.RecordCount <> 0 Then
        ProgressBar1.Max = RstTmp2.RecordCount
    
        RstTmp2.MoveFirst
        
        RstTmp2.Sort = "registro"
        
        Do While Not RstTmp2.EOF
            DoEvents
            ProgressBar1.Value = ProgressBar1.Value + 1
            '--SI SE NTERRUMPE EL PROCESO => SALIR
            If BAND_INTERRUMPIR = True Then GoTo SALIR
            '-----------------------------------------------
            Fg2.Rows = Fg2.Rows + 1
            mCol = 0
            For mColCampo = 2 To RstTmp2.Fields.Count - 1
                mCol = mCol + 1
                Select Case LCase(RstTmp2.Fields(mColCampo).Name)
                    Case "libdesc", "registro", "registroref", "glosa", "numruc", "apenom", "tdocdesc", "docsustenta", "ctanum", "ctadesc", "simbolo"
                        Fg2.TextMatrix(Fg2.Rows - 1, mCol) = NulosC(RstTmp2.Fields(mColCampo))
                    Case "fchdoc", "fchope"
                        Fg2.TextMatrix(Fg2.Rows - 1, mCol) = Format(RstTmp2.Fields(mColCampo), FORMAT_DATE)
                    Case "tc", "tipcam"
                        Fg2.TextMatrix(Fg2.Rows - 1, mCol) = Format(RstTmp2.Fields(mColCampo), "0.000")
                    Case "debe"
                        Fg2.TextMatrix(Fg2.Rows - 1, mCol) = Format(RstTmp2.Fields(mColCampo), FORMAT_MONTO)
                        'xTotal1 = xTotal1 + NulosN(RstTmp2("debe"))
                        xTotal1 = xTotal1 + NulosN(Fg2.TextMatrix(Fg2.Rows - 1, mCol))
                    Case "haber"
                        Fg2.TextMatrix(Fg2.Rows - 1, mCol) = Format(RstTmp2.Fields(mColCampo), FORMAT_MONTO)
                        'xTotal2 = xTotal2 + NulosN(RstTmp2("haber"))
                        xTotal2 = xTotal2 + NulosN(Fg2.TextMatrix(Fg2.Rows - 1, mCol))
                    Case "saldo"
'                        xSaldo = xSaldo + NulosN(RstTmp2("saldo"))
'                        Fg2.TextMatrix(Fg2.Rows - 1, mCol) = Format(xSaldo, FORMAT_MONTO)
                        
                                If UCase(RstTmp2.Fields("tipsal") & "") = "D" Or NulosC(RstTmp2.Fields("tipsal")) = "" Then
                                    xSaldo = xSaldo + Format((NulosN(RstTmp2(mColDebe + 1)) - NulosN(RstTmp2(mColHaber + 1))), FORMAT_MONTO)
                                Else
                                    xSaldo = xSaldo + Format((NulosN(RstTmp2(mColHaber + 1)) - NulosN(RstTmp2(mColDebe + 1))), FORMAT_MONTO)
                                End If
                                
                                Fg2.TextMatrix(Fg2.Rows - 1, mCol) = Format(xSaldo, FORMAT_MONTO)
                        
                        
                    Case Else
                        Fg2.TextMatrix(Fg2.Rows - 1, mCol) = NulosC(RstTmp2.Fields(mColCampo))
                End Select
                
            Next mColCampo
            
            RstTmp2.MoveNext
            If RstTmp2.EOF = True Then
                Fg2.Rows = Fg2.Rows + 1
                Fg2.TextMatrix(Fg2.Rows - 1, mColDebe - 1) = "Total =>"
                Fg2.TextMatrix(Fg2.Rows - 1, mColDebe) = Format(xTotal1, FORMAT_MONTO)
                Fg2.TextMatrix(Fg2.Rows - 1, mColHaber) = Format(xTotal2, FORMAT_MONTO)
                
                FORMATO_CELDA Fg2, Fg2.Rows - 1, mColDebe - 1, , True
                FORMATO_CELDA Fg2, Fg2.Rows - 1, mColDebe, , True
                FORMATO_CELDA Fg2, Fg2.Rows - 1, mColHaber, , True
                
                Fg2.Rows = Fg2.Rows + 1
                Exit Do
            End If
        Loop
        
    End If
    
    
    '--ajustando las columnas de acuerdo a los importes
    Fg2.AutoSizeMode = flexAutoSizeColWidth
    Fg2.AutoSize mColDebe
    Fg2.AutoSize mColHaber
    Fg2.AutoSize mColSaldo
        
SALIR:
    Frame11.Visible = False
    Me.MousePointer = vbDefault
End Sub

'*******************************************************************************************

Private Sub pBuscarAsiento()
    Dim xfrm As New SGI2_funciones.formularios
    xfrm.AsientoBuscar xCon
    Set xfrm = Nothing
End Sub




'***********************************************************************************************
'***********************************************************************************************


'Private Sub pImprimir()
'
'    On Error GoTo error
'
'
'    Me.MousePointer = vbHourglass
'
'    Dim X_PRINT As New SGI2_funciones.formularios
'    Dim xMoneda As String
'    Dim nPeriodo  As String
'    If opt_fecha(0).Value = True Then  '--por fecha
'        If CDate(Me.TxtFchIni.Valor) <> CDate(Me.TxtFchFin.Valor) Then
'            nPeriodo = "Del: " + CStr(TxtFchIni.Valor) + " Al: " + CStr(TxtFchFin.Valor)
'        Else
'            nPeriodo = "Al: " + CStr(TxtFchIni.Valor)
'        End If
'    Else '--por periodo
'        If mMesIni = mMesFin Then
'            nPeriodo = "Periodo : " + cmd_periodo1.Caption
'        Else
'            nPeriodo = "Periodo : De " + cmd_periodo1.Caption & " A " & cmd_periodo2.Caption
'        End If
'    End If
'
'    xMoneda = LblMoneda.Caption
'
'    X_PRINT.Imprimir_x_VSFlexGrid fg1(TabOne1.CurrTab), "HOJA DE TRABAJO ", "(Expresado en " + xMoneda + ")", nPeriodo, False, True
'    Set X_PRINT = Nothing
'
'    Me.MousePointer = vbDefault
'    Exit Sub
'error:
'    Me.MousePointer = vbDefault
'    SHOW_ERROR Me.Name, "CmdImprimir_Click"
'End Sub





Private Function fValidarConsulta() As Boolean
    '--FUNCION QUE VALIDARA LA CONSULTA
    If AnoTra = "" Then
        MsgBox "No Hay Año de trabajo" + vbCr + "No puede continuar", vbExclamation, xTitulo
        Exit Function
    End If
    
    If opt_fecha(0).Value = True Then '--por fecha
        If NulosC(TxtFchIni.Valor) = "" Then
            MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchIni.SetFocus
            Exit Function
        End If
        
        If NulosC(TxtFchFin.Valor) = "" Then
            MsgBox "No ha especificado la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchFin.SetFocus
            Exit Function
        End If
        
        If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
            MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchIni.SetFocus
            Exit Function
        End If
        
        If (Year(TxtFchIni.Valor) <> Year(TxtFchFin.Valor)) Then
            MsgBox "El rango de fechas debe estar en el Año de Trabajo : " + CStr(AnoTra), vbExclamation, xTitulo
            TxtFchIni.SetFocus
            Exit Function
        ElseIf Year(TxtFchIni.Valor) <> CStr(AnoTra) Then
            MsgBox "El rango de fechas debe estar en el Año de Trabajo : " + CStr(AnoTra), vbExclamation, xTitulo
            TxtFchIni.SetFocus
            Exit Function
        End If
        
    Else '--por periodo
    
        If mMesIni > mMesFin Then
            MsgBox "El periodo de inicio debe ser inferior o igual al periodo final", vbExclamation, xTitulo
            cmd_periodo1.SetFocus
            Exit Function
        End If
        
    End If
    
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "Seleccione la Moneda", vbExclamation, xTitulo
        Exit Function
    End If
    
    
    fValidarConsulta = True
End Function

Private Sub pImprimir()
    Dim xMoneda As String
    Dim nPeriodo As String
    
    If opt_fecha(0).Value = True Then  '--por fecha
        If CDate(Me.TxtFchIni.Valor) <> CDate(Me.TxtFchFin.Valor) Then
            nPeriodo = "Del: " + CStr(TxtFchIni.Valor) + " Al: " + CStr(TxtFchFin.Valor)
        Else
            nPeriodo = "Al: " + CStr(TxtFchIni.Valor)
        End If
    Else '--por periodo
        If mMesIni = mMesFin Then
            nPeriodo = "Periodo : " & LblPerIni.Caption
        Else
            nPeriodo = "Periodo : De " + LblPerIni.Caption & " A " & LblPerFin.Caption
        End If
    End If
    
    xMoneda = LblMoneda.Caption
    
    Dim RstTmp As New ADODB.Recordset
    Dim A As Integer
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT con_formatostipodet.*,con_formatostipo.rpttitulo, con_formatostipo.rpttamdet, con_formatostipo.rpttamcab " _
        & " FROM con_formatostipodet INNER JOIN con_formatostipo ON (con_formatostipo.id = con_formatostipodet.idformatotipo) AND (con_formatostipodet.idformato = con_formatostipo.idformato) " _
        & " WHERE (((con_formatostipo.idformato)=7) AND ((con_formatostipodet.mostrar)=-1) AND ((con_formatostipo.defecto)=-1)) " _
        & " ORDER BY con_formatostipodet.orden", xCon
    
    Dim xCampos() As String
    Dim xFil, xCol As Double
    Dim xIndice As Integer
    
    xIndice = TabOne1.CurrTab
    
    ReDim xCampos(Fg1(xIndice).Rows - 2, Fg1(xIndice).Cols - 1)
    
    Dim xFila As Double
    xFila = 0
    For xFil = 1 To Fg1(xIndice).Rows - 1
        For xCol = 1 To Fg1(xIndice).Cols - 1
            xCampos(xFila, xCol) = Fg1(xIndice).TextMatrix(xFil, xCol)
        Next xCol
        xFila = xFila + 1
    Next xFil
    
    Rst.MoveFirst
    For A = 1 To Rst.RecordCount
        If NulosC(xCampos(0, A)) = NulosC(Rst("abrev")) Then
            If Rst("imprimir") = False Then
                xCampos(0, A) = ""
            End If
        End If
        Rst.MoveNext
        If Rst.EOF = True Then Exit For
    Next A
    
    Rst.MoveFirst
    
    Dim xfrm As New eps_librerias.IMPRIMIR
    
    xfrm.Cabecera1 = NomEmp
    xfrm.Cabecera2 = "RUC Nº: " & NumRUC
    xfrm.Fecha = Format(Date, "dd/mm/yyyy")
    xfrm.Titulo1 = NulosC(Rst("rpttitulo")) & " (Expresado en " & xMoneda & ")"
    xfrm.Titulo2 = nPeriodo
    xfrm.TamañoFuente = NulosN(Rst("rpttamdet"))
    xfrm.TamañoCabecera = NulosN(Rst("rpttamcab"))
    xfrm.FuenteCabecera = "Courier New"
    xfrm.Posicion_Hoja = Horizontal
    xfrm.Tamaño_Hoja = A_4
    xfrm.TextoConsiderar = " "
    xfrm.TextoConsiderarAncho = 1
    xfrm.ImprimirArray xCampos, Rst
    Set xfrm = Nothing
    Set Rst = Nothing
    
End Sub


Sub Configurar()
    Dim xform As New SGI2_funciones.Varias
    If xform.CambioOpcionLiro(7, xCon, 1) = True Then
        
    End If
    Set xform = Nothing
End Sub

Private Sub pConsultar()
    If fValidarConsulta() = False Then Exit Sub
    
    FraDetalle.Visible = False
    
    If TabOne1.CurrTab <> 3 Then
        Procesar1 TabOne1.CurrTab
    Else
        CargarAnalisis
    End If
    
End Sub



Private Sub CargarAnalisis()


    '===================================================================================================
    'creado: 08/08/09 Johan Castro
    'Propósito: Mostrar la información del balance de comprobacion en periodos
    '
    'Entradas:  Fecha Incial, Fecha Final
    '
    'Resultados: balance de comprobacion en pantalla por periodo
    '
    'Modificado: 15/12/10 Johan Castro
    '            Agregar funcion nulosc() cuando escribe datos de cuenta y nombre de cuenta
    '===================================================================================================
    
    'LEYENDA:
    'SI: Saldos Iniciales
    'MP: Movimientos del Periodo
    'SM: Sumas del Mayor
    'SA: Saldos Finales
    'CB: Cuentas de Balance
    'CT: Cuentas de Transferencia
    'GN: Ganancias por Naturaleza
    'GF: Ganancias por Funcion


    Dim nSQL As String
    Dim Rst As New ADODB.Recordset
    Dim nSQLIdLibro As String
    Dim nSQLWhere As String
    Dim nSQLTipoPersona As String
    Dim nSQLAjuste  As String
    Dim nSQLCierre As String
    
    Dim FchIni, FchFin As Date
    Dim A, C As Integer '--cantidad de meses a consultar
    Dim mRow As Long
    
    '--dar formato al grid
    SeteaAnalisis
    '------------------------------
    Frame11.Left = 3120
    Frame11.Top = 3930
    Frame11.Visible = True
    ProgressBar1.Min = 0
    ProgressBar1.Value = 0

    
    
    If NulosN(LblIdLibro.Caption) <> 0 Then nSQLIdLibro = " and con_diario.idlib=" & LblIdLibro
    '--si selecciona un tipo de persona
    If lbl_cod(1).Caption <> "" Then
        If NulosN(lbl_cod(0).Caption) <> 0 Then
            nSQLTipoPersona = " and con_diario.ridtipper = " & lbl_cod(0).Caption & " and con_diario.ridper = " & lbl_cod(1).Caption
        Else
            '--buscar
            nSQLTipoPersona = " and " & Replace(lbl_cod(1).Caption, "cod", "con_diario.ridper")
        End If
    End If
    
    '--para ajuste por diferencia de cambio
    nSQLAjuste = " AND (con_diario.ajuste in (0, " & NulosN(TxtIdMon.Text) & ") ) "
    '-----------------------------------------------
    nSQLCierre = " AND (con_diario.idmes<>13) "
    '-----------------------------------------------
    
    
    Fg1(3).Rows = Fg1(3).FixedRows
    DoEvents
    C = 0
    
    For A = Month(TxtFchIni.Valor) To Month(TxtFchFin.Valor)
    
        Label6.Caption = "Procesando Consulta - " & NomMes(CInt(A))
        
        DoEvents
        
        FchIni = CDate("01/" & Format(A, "00") & "/" & Year(TxtFchIni.Valor))
        FchFin = CDate(HallaDiasMes(CDate(FchIni)) & "/" & Format(A, "00") & "/" & Year(TxtFchIni.Valor))

        '--19/04/09
        '--se cambia los saldos iniciales solo debera de mostrar debe o harer
        '-- IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol) AS SIDebSol,
        '-- IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol) AS SIHabSol,
        nSQL = ""
        nSQL = "SELECT con_planctas.cuenta, con_planctas.descripcion, " _
            + vbCr + " IIf(((IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))-(IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol)))>0,((IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))-(IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))),0) AS SIDebSol, " _
            + vbCr + " IIf(((IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))-(IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol)))>0,((IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))-(IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))),0) AS SIHabSol, " _
            + vbCr + " IIf(MovPeriodo.DebSol Is Null,0,MovPeriodo.DebSol) AS MPDebSol, " _
            + vbCr + " IIf(MovPeriodo.HabSol Is Null,0,MovPeriodo.HabSol) AS MPHabSol, " _
            + vbCr + " [SIDebSol]+[MPDebSol] AS SMDebSol, " _
            + vbCr + " [SIHabSol]+[MPHabSol] AS SMHabSol, " _
            + vbCr + " IIf((SMDebSol-SMHabSol)>0,(SMDebSol-SMHabSol),0) AS SADebSol, " _
            + vbCr + " IIf((SMHabSol-SMDebSol)>0,(SMHabSol-SMDebSol),0) AS SAHabSol, " _
            + vbCr + " IIf(con_planctas.iddes=1 Or con_planctas.iddes2=1,SADebSol,0) AS CBDebSol, " _
            + vbCr + " IIf(con_planctas.iddes=1 Or con_planctas.iddes2=1,SAHabSol,0) AS CBHabSol, " _
            + vbCr + " IIf(con_planctas.iddes=4 Or con_planctas.iddes2=4,SADebSol,0) AS CTDebSol, " _
            + vbCr + " IIf(con_planctas.iddes=4 Or con_planctas.iddes2=4,SAHabSol,0) AS CTHabSol, " _
            + vbCr + " IIf(con_planctas.iddes=2 Or con_planctas.iddes2=2,SADebSol,0) AS GNDebSol, " _
            + vbCr + " IIf(con_planctas.iddes=2 Or con_planctas.iddes2=2,SAHabSol,0) AS GNHabSol, " _
            + vbCr + " IIf(con_planctas.iddes=3 Or con_planctas.iddes2=3,SADebSol,0) AS GFDebSol, " _
            + vbCr + " IIf(con_planctas.iddes=3 Or con_planctas.iddes2=3,SAHabSol,0) As GFHabSol, " _
            + vbCr + " con_planctas.id AS IdCta "
        
        If NulosN(TxtIdMon.Text) = 2 Then
            nSQL = Replace(nSQL, "DebSol", "DebDol")
            nSQL = Replace(nSQL, "HabSol", "HabDol")
        End If
        
        '--movimientos del periodo
        nSQL = nSQL _
            + vbCr + " FROM (con_planctas LEFT JOIN " _
            + vbCr + " ( " _
            + vbCr + " SELECT con_planctas.id AS IdCta, con_planctas.cuenta, con_planctas.descripcion, " _
            + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS DebSol, " _
            + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS HabSol, " _
            + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS DebDol, " _
            + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabdol)) As HabDol " _
            + vbCr + " FROM (con_planctas RIGHT JOIN con_diario ON con_planctas.id=con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc=con_tc.fecha " _
            + vbCr + " WHERE (((con_diario.fchasi) Between CDate('" & FchIni & "') And CDate('" & FchFin & "'))) " & nSQLIdLibro & nSQLTipoPersona & nSQLAjuste _
            + vbCr + " GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion " _
            + vbCr + " ORDER BY con_planctas.cuenta " _
            + vbCr + " ) AS MovPeriodo ON con_planctas.id = MovPeriodo.IdCta) " _
            + vbCr + " Left Join "
        
        '--saldos iniciales
        nSQL = nSQL _
            + vbCr + " ( " _
            + vbCr + " SELECT con_planctas.id AS IdCta, con_planctas.cuenta, con_planctas.descripcion, " _
            + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS DebSol, " _
            + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS HabSol, " _
            + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS DebDol, " _
            + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabdol)) As HabDol " _
            + vbCr + " FROM (con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
            + vbCr + " WHERE  ((con_diario.fchasi) Is Null Or (con_diario.fchasi)<CDate('" & FchIni & "')) " & nSQLIdLibro & nSQLTipoPersona & nSQLAjuste & nSQLCierre _
            + vbCr + " GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion " _
            + vbCr + " ORDER BY con_planctas.cuenta " _
            + vbCr + " ) AS SaldosIni "
        
        nSQLWhere = nSQLIdLibro & nSQLTipoPersona & nSQLAjuste & " AND (  (((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')))  OR  (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "')  OR  (con_diario.fchasi) is null  )"
    
    
        nSQL = nSQL _
            + vbCr + " ON con_planctas.id = SaldosIni.IdCta " _
            + vbCr + " WHERE con_planctas.id In (SELECT con_diario.idcue FROM con_diario " & IIf(nSQLWhere <> "", "WHERE " & Mid(nSQLWhere, 5), "") & "   ) " _
            + vbCr + " ORDER BY con_planctas.cuenta; "

        '--si seleccionar por periodo
        If opt_fecha(1).Value = True Then
            '--movimiento del periodo
            nSQL = Replace(nSQL, "(((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')))", "( con_diario.idmes>=" & mMesIni & " And con_diario.idmes <= " & mMesFin & " )")
            '--saldos iniciales
            nSQL = Replace(nSQL, "((con_diario.fchasi) Is Null Or (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "'))", "con_diario.idmes < " & mMesIni)
            '--filtro para where
            nSQL = Replace(nSQL, "OR  (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "')  OR  (con_diario.fchasi) is null ", " OR con_diario.idmes <= " & mMesIni)
            
        End If

        
        Me.MousePointer = vbHourglass
        DoEvents
        Set Rst = Nothing
        RST_Busq Rst, nSQL, xCon
        Dim mCol&
        If Rst.RecordCount <> 0 Then ProgressBar1.Max = Rst.RecordCount + 1
        
        ProgressBar1.Value = 1
        
        If C = 0 Then Fg1(3).Rows = Rst.RecordCount + 1 + Fg1(3).FixedRows
        
        mRow = Fg1(3).FixedRows
       
        '--proceder a cargar los datos
        Do While Not Rst.EOF
            DoEvents
            '-----------------------------------------------
            ProgressBar1.Value = ProgressBar1.Value + 1
            DoEvents
            '--SI SE NTERRUMPE EL PROCESO => SALIR
            If BAND_INTERRUMPIR = True Then GoTo SALIR
            DoEvents
            For mCol = 0 To Rst.Fields.Count - 1
                
                If mCol <= 1 Then
                    Fg1(3).TextMatrix(mRow, 20 * C + mCol + 1) = NulosC(Rst(mCol))
                    
                ElseIf mCol < 18 Then
                    '--importes
                    Fg1(3).TextMatrix(mRow, 20 * C + mCol + 1) = Format(NulosN(Rst(mCol)), FORMAT_MONTO)
                    
                Else
                    '--idcta
                    Fg1(3).TextMatrix(mRow, 20 * C + mCol + 1) = NulosN(Rst(mCol))
                
                End If
            Next mCol
            mRow = mRow + 1
            Rst.MoveNext
            
        Loop
        C = C + 1
    Next A
    '--colocar los resumenes
    TotalizarAnalisis
    
SALIR:
    Frame11.Visible = False
    Me.MousePointer = vbDefault

End Sub


Sub SeteaAnalisis()
    'usamos la columna 19 para almacenar el destino de cada cuenta en la hoja de trabajo
    Dim A As Integer
    Dim B As Integer
    Dim C As Integer
    Dim x As Integer
    
    A = 3
    
    Fg1(A).Cols = 1
    Fg1(A).Rows = 3
    Fg1(A).FixedRows = 3
    
    C = 0
    
    For x = Month(TxtFchIni.Valor) To Month(TxtFchFin.Valor)
        
         Fg1(A).Cols = Fg1(A).Cols + 20
         
         Fg1(A).TextMatrix(0, 20 * C + 1) = "          1"
         Fg1(A).TextMatrix(1, 20 * C + 1) = "          1"
         Fg1(A).TextMatrix(2, 20 * C + 1) = "          1"
         Fg1(A).TextMatrix(0, 20 * C + 1) = "Nº Cuenta"
         Fg1(A).TextMatrix(1, 20 * C + 1) = "Nº Cuenta"
         Fg1(A).TextMatrix(2, 20 * C + 1) = "Nº Cuenta"
         Fg1(A).TextMatrix(0, 20 * C + 2) = "Descripción"
         Fg1(A).TextMatrix(1, 20 * C + 2) = "Descripción"
         Fg1(A).TextMatrix(2, 20 * C + 2) = "Descripción"
         
         'Fg1.MergeCells = flexMergeFree
         Fg1(A).Redraw = False
         Fg1(A).MergeCol(0) = True
         Fg1(A).MergeCol(20 * C + 1) = True
         Fg1(A).MergeCol(20 * C + 2) = True
         
         Fg1(A).ColWidth(20 * C + 19) = 0
         Fg1(A).ColWidth(20 * C + 20) = 0
         
         Fg1(A).MergeCells = 2
         Fg1(A).Redraw = True
         
         With Fg1(A)
             .MergeCells = flexMergeFree
             .MergeRow(-1) = True
             
             .Cell(flexcpText, 1, 20 * C + 3, 1, 20 * C + 4) = "Saldos Iniciales"
             .Cell(flexcpText, 1, 20 * C + 5, 1, 20 * C + 6) = "Movimiento del Periodo"
             .Cell(flexcpText, 1, 20 * C + 7, 1, 20 * C + 8) = "Sumas del Mayor"
             .Cell(flexcpText, 1, 20 * C + 9, 1, 20 * C + 10) = "Saldos Finales" 'a " & NomMes(x)
             .Cell(flexcpText, 1, 20 * C + 11, 1, 20 * C + 12) = "Cuentas del Balance"
             .Cell(flexcpText, 1, 20 * C + 13, 1, 20 * C + 14) = "Transferencias y Canc."
             .Cell(flexcpText, 1, 20 * C + 15, 1, 20 * C + 16) = "Resultados x Naturaleza"
             .Cell(flexcpText, 1, 20 * C + 17, 1, 20 * C + 18) = "Resultados x Función"

             .Row = 1: .Col = 20 * C + 3: .CellAlignment = flexAlignCenterCenter
             .Row = 1: .Col = 20 * C + 5: .CellAlignment = flexAlignCenterCenter
             .Row = 1: .Col = 20 * C + 7: .CellAlignment = flexAlignCenterCenter
             .Row = 1: .Col = 20 * C + 9: .CellAlignment = flexAlignCenterCenter
             .Row = 1: .Col = 20 * C + 11: .CellAlignment = flexAlignCenterCenter
             .Row = 1: .Col = 20 * C + 13: .CellAlignment = flexAlignCenterCenter
             .Row = 1: .Col = 20 * C + 15: .CellAlignment = flexAlignCenterCenter
             .Row = 1: .Col = 20 * C + 17: .CellAlignment = flexAlignCenterCenter
             
             .Cell(flexcpBackColor, 0, 0, Fg1(A).Rows - 1, Fg1(A).Cols - 1) = &H8000000F
             
         End With

         Fg1(A).TextMatrix(2, 20 * C + 19) = "IdCta"
         
         '--nombre del periodo
         GRID_COMBINAR Fg1(3), 0, 20 * C + 3, 0, 20 * C + 19, NomMes(x), , True
         '-----------------------------------------------------------------------
         
         For B = 3 To 18
         
            Fg1(A).ColWidth(20 * C + B) = 1100
            If B Mod 2 = 0 Then
                Fg1(A).TextMatrix(2, 20 * C + B) = "Haber"
            Else
                Fg1(A).TextMatrix(2, 20 * C + B) = "Debe"
            End If
             
             Fg1(A).ColAlignment(20 * C + B) = flexAlignRightCenter
         
         Next B
        
        '--
        C = C + 1
        
        If C > 1 Then
            Fg1(A).ColWidth(20 * (C - 1) + 1) = 0
            Fg1(A).ColWidth(20 * (C - 1) + 2) = 0
        End If
        
        
        For B = 0 To chk.Count - 1
            If chk(B).Value = 0 Then
                
                Fg1(A).ColWidth(20 * (C - 1) + 3 + (2 * B)) = 0
                Fg1(A).ColWidth(20 * (C - 1) + 4 + (2 * B)) = 0
                
            End If
        Next B
        
    Next x
    
    '---------------------------------------------
    Fg1(3).ColWidth(1) = 1035
    Fg1(3).ColWidth(2) = 2670
    Fg1(3).ColAlignment(1) = flexAlignLeftCenter
    Fg1(3).FrozenCols = 2
    
End Sub


Private Sub TotalizarAnalisis()
    Dim A As Integer
    Dim C As Integer
    Dim x As Integer
    
    Fg1(3).Rows = Fg1(3).Rows + 1
    Fg1(3).Rows = Fg1(3).Rows + 1
    Fg1(3).Rows = Fg1(3).Rows + 1
    
    Fg1(3).AutoSizeMode = flexAutoSizeColWidth
    
    C = 0
    
    For x = Month(TxtFchIni.Valor) To Month(TxtFchFin.Valor)
    
        Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 2) = "T O T A L E S ==>"
        
        '-totalizando
        For A = 3 To 18
            Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + A) = Format(GRID_SUMAR_COL(Fg1(3), 20 * C + A), FORMAT_MONTO)
            '--ajustando las columnas de acuerdo a los importes
            If Fg1(3).ColWidth(20 * C + A) <> 0 Then
                Fg1(3).AutoSize 20 * C + A
                If Fg1(3).ColWidth(20 * C + A) < 850 Then Fg1(3).ColWidth(20 * C + A) = 880
            End If
        Next A
        '-------------------------
        
        
        Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 2) = "R E S U L T A D O ==>"
            
        If NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 10)) - NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 9)) > 0 Then
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 9) = NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 9)) - NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 10))
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 9) = Format(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 9), FORMAT_MONTO)
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 10) = "0.00"
        Else
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 9) = "0.00"
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 10) = NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 10)) - NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 9))
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 10) = Format(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 10), FORMAT_MONTO)
        End If
    
        If NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 12)) - NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 11)) > 0 Then
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 11) = NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 12)) - NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 11))
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 11) = Format(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 11), FORMAT_MONTO)
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 12) = "0.00"
        Else
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 11) = "0.00"
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 12) = NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 11)) - NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 12))
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 12) = Format(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 12), FORMAT_MONTO)
        End If
        
        If NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 14)) - NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 13)) > 0 Then
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 13) = NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 14)) - NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 13))
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 13) = Format(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 13), FORMAT_MONTO)
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 14) = "0.00"
        Else
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 13) = "0.00"
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 14) = NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 13)) - NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 14))
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 14) = Format(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 14), FORMAT_MONTO)
        End If
        
        If NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 16)) - NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 15)) > 0 Then
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 15) = NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 16)) - NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 15))
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 15) = Format(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 15), FORMAT_MONTO)
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 16) = "0.00"
        Else
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 15) = "0.00"
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 16) = NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 15)) - NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 16))
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 16) = Format(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 16), FORMAT_MONTO)
        End If
        
        If NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 18)) - NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 17)) > 0 Then
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 17) = NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 18)) - NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 17))
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 17) = Format(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 17), FORMAT_MONTO)
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 18) = "0.00"
        Else
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 17) = "0.00"
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 18) = NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 17)) - NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 18))
            Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 18) = Format(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 18), FORMAT_MONTO)
        End If
        
        Fg1(3).TextMatrix(Fg1(3).Rows - 1, 20 * C + 2) = "S U M A S  I G U A L E S ==>"
        
        Fg1(3).TextMatrix(Fg1(3).Rows - 1, 20 * C + 9) = Format(NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 9)) + NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 9)), FORMAT_MONTO)
        Fg1(3).TextMatrix(Fg1(3).Rows - 1, 20 * C + 10) = Format(NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 10)) + NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 10)), FORMAT_MONTO)
        
        Fg1(3).TextMatrix(Fg1(3).Rows - 1, 20 * C + 11) = Format(NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 11)) + NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 11)), FORMAT_MONTO)
        Fg1(3).TextMatrix(Fg1(3).Rows - 1, 20 * C + 12) = Format(NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 12)) + NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 12)), FORMAT_MONTO)
    
        Fg1(3).TextMatrix(Fg1(3).Rows - 1, 20 * C + 13) = Format(NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 13)) + NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 13)), FORMAT_MONTO)
        Fg1(3).TextMatrix(Fg1(3).Rows - 1, 20 * C + 14) = Format(NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 14)) + NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 14)), FORMAT_MONTO)
    
        Fg1(3).TextMatrix(Fg1(3).Rows - 1, 20 * C + 15) = Format(NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 15)) + NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 15)), FORMAT_MONTO)
        Fg1(3).TextMatrix(Fg1(3).Rows - 1, 20 * C + 16) = Format(NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 16)) + NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 16)), FORMAT_MONTO)
    
        Fg1(3).TextMatrix(Fg1(3).Rows - 1, 20 * C + 17) = Format(NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 17)) + NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 17)), FORMAT_MONTO)
        Fg1(3).TextMatrix(Fg1(3).Rows - 1, 20 * C + 18) = Format(NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 3, 20 * C + 18)) + NulosN(Fg1(3).TextMatrix(Fg1(3).Rows - 2, 20 * C + 18)), FORMAT_MONTO)
    
        '*********************************************************************************************************
        '--color de las columnas
        GRID_COLOR_FONDO Fg1(3), Fg1(3).FixedRows, 20 * C + 1, Fg1(3).Rows - 1, 20 * C + 2, &HD5EAFF '&HECFFFF
        GRID_COLOR_FONDO Fg1(3), Fg1(3).FixedRows, 20 * C + 5, Fg1(3).Rows - 1, 20 * C + 6, &HD5EAFF
        GRID_COLOR_FONDO Fg1(3), Fg1(3).FixedRows, 20 * C + 9, Fg1(3).Rows - 1, 20 * C + 10, &HD5EAFF
        GRID_COLOR_FONDO Fg1(3), Fg1(3).FixedRows, 20 * C + 13, Fg1(3).Rows - 1, 20 * C + 14, &HD5EAFF
        GRID_COLOR_FONDO Fg1(3), Fg1(3).FixedRows, 20 * C + 17, Fg1(3).Rows - 1, 20 * C + 18, &HD5EAFF
        '--color de fondo
        GRID_COLOR_FONDO Fg1(3), Fg1(3).Rows - 3, 20 * C + 1, Fg1(3).Rows - 1, Fg1(3).Cols - 1, &HFDF5E8 '&HE0FEE7
        
        '*********************************************************************************************************
        
        C = C + 1
        
    Next x
    

End Sub



Private Sub Cargar1_XXXXX(Indice As Integer)
    '===================================================================================================
    'creado: 25/12/08
    'Propósito: Mostrar la información del balance de comprobacion
    '
    'Entradas:  Indice = Tipo de Reporte
    '
    'Resultados: balance de comprobacion libros en pantalla
    '===================================================================================================
    'LEYENDA:
    'SI: Saldos Iniciales
    'MP: Movimientos del Periodo
    'SM: Sumas del Mayor
    'SA: Saldos Al
    'CB: Cuentas de Balance
    'CT: Cuentas de Transferencia
    'GN: Ganancias por Naturaleza
    'GF: Ganancias por Funcion


    Dim nSQL As String
    Dim Rst As New ADODB.Recordset
    Dim nSQLIdLibro As String
    Dim nSQLTipoPersona As String
    Dim nSQLAjuste  As String
    
    Dim FchIni, FchFin As Date
    
    
    
    
    Frame11.Left = 3120
    Frame11.Top = 3930
    Frame11.Visible = True
    ProgressBar1.Min = 0
    ProgressBar1.Value = 0

    
    If NulosN(LblIdLibro.Caption) <> 0 Then nSQLIdLibro = " and con_diario.idlib=" & LblIdLibro
    '--si selecciona un tipo de persona
    If lbl_cod(1).Caption <> "" Then
        If NulosN(lbl_cod(0).Caption) <> 0 Then
            nSQLTipoPersona = " and con_diario.ridtipper = " & lbl_cod(0).Caption & " and con_diario.ridper = " & lbl_cod(1).Caption
        Else
            '--buscar
            nSQLTipoPersona = " and " & Replace(lbl_cod(1).Caption, "cod", "con_diario.ridper")
        End If
    End If
    
        '--para ajuste por diferencia de cambio
    nSQLAjuste = " AND (con_diario.ajuste in (0, " & NulosN(TxtIdMon.Text) & ") ) "
    '-----------------------------------------------

    
    Fg1(0).Rows = Fg1(0).FixedRows
    DoEvents
    
    '--19/04/09
    '--se cambia los saldos iniciales solo debera de mostrar debe o haber
    '-- IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol) AS SIDebSol,
    '-- IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol) AS SIHabSol,
    nSQL = "SELECT con_planctas.cuenta, con_planctas.descripcion, " _
        + vbCr + " IIf(((IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))-(IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol)))>0,((IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))-(IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))),0) AS SIDebSol, " _
        + vbCr + " IIf(((IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))-(IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol)))>0,((IIf(SaldosIni.HabSol Is Null,0,SaldosIni.HabSol))-(IIf(SaldosIni.DebSol Is Null,0,SaldosIni.DebSol))),0) AS SIHabSol, " _
        + vbCr + " IIf(MovPeriodo.DebSol Is Null,0,MovPeriodo.DebSol) AS MPDebSol, " _
        + vbCr + " IIf(MovPeriodo.HabSol Is Null,0,MovPeriodo.HabSol) AS MPHabSol, " _
        + vbCr + " [SIDebSol]+[MPDebSol] AS SMDebSol, " _
        + vbCr + " [SIHabSol]+[MPHabSol] AS SMHabSol, " _
        + vbCr + " IIf((SMDebSol-SMHabSol)>0,(SMDebSol-SMHabSol),0) AS SADebSol, " _
        + vbCr + " IIf((SMHabSol-SMDebSol)>0,(SMHabSol-SMDebSol),0) AS SAHabSol, " _
        + vbCr + " IIf(con_planctas.iddes=1 Or con_planctas.iddes2=1,SADebSol,0) AS CBDebSol, " _
        + vbCr + " IIf(con_planctas.iddes=1 Or con_planctas.iddes2=1,SAHabSol,0) AS CBHabSol, " _
        + vbCr + " IIf(con_planctas.iddes=4 Or con_planctas.iddes2=4,SADebSol,0) AS CTDebSol, " _
        + vbCr + " IIf(con_planctas.iddes=4 Or con_planctas.iddes2=4,SAHabSol,0) AS CTHabSol, " _
        + vbCr + " IIf(con_planctas.iddes=2 Or con_planctas.iddes2=2,SADebSol,0) AS GNDebSol, " _
        + vbCr + " IIf(con_planctas.iddes=2 Or con_planctas.iddes2=2,SAHabSol,0) AS GNHabSol, " _
        + vbCr + " IIf(con_planctas.iddes=3 Or con_planctas.iddes2=3,SADebSol,0) AS GFDebSol, " _
        + vbCr + " IIf(con_planctas.iddes=3 Or con_planctas.iddes2=3,SAHabSol,0) As GFHabSol, " _
        + vbCr + " con_planctas.iddes,con_planctas.iddes2,con_planctas.id AS IdCta "
    
    If NulosN(TxtIdMon.Text) = 2 Then
        nSQL = Replace(nSQL, "DebSol", "DebDol")
        nSQL = Replace(nSQL, "HabSol", "HabDol")
    End If
    '--movimientos del periodo
    nSQL = nSQL _
        + vbCr + " FROM (con_planctas LEFT JOIN " _
        + vbCr + " ( " _
        + vbCr + " SELECT con_planctas.id AS IdCta, con_planctas.cuenta, con_planctas.descripcion, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS DebSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS HabSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS DebDol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabdol)) As HabDol " _
        + vbCr + " FROM (con_planctas RIGHT JOIN con_diario ON con_planctas.id=con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc=con_tc.fecha " _
        + vbCr + " WHERE (((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) " & nSQLIdLibro & nSQLTipoPersona & nSQLAjuste _
        + vbCr + " GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion " _
        + vbCr + " ORDER BY con_planctas.cuenta " _
        + vbCr + " ) AS MovPeriodo ON con_planctas.id = MovPeriodo.IdCta) " _
        + vbCr + " Left Join "
    
    '--saldos iniciales
    nSQL = nSQL _
        + vbCr + " ( " _
        + vbCr + " SELECT con_planctas.id AS IdCta, con_planctas.cuenta, con_planctas.descripcion, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebdol=0,0,con_diario.impdebdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebsol)) AS DebSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=2,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabdol=0,0,con_diario.imphabdol*(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabsol)) AS HabSol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.impdebsol=0,0,con_diario.impdebsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.impdebdol)) AS DebDol, " _
        + vbCr + " Sum(IIf(con_diario.idmon=1,IIf(iif( con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven))=0 or con_diario.imphabsol=0,0,con_diario.imphabsol/(iif( con_diario.aplicatc=-1, con_diario.tc,con_tc.impven))),con_diario.imphabdol)) As HabDol " _
        + vbCr + " FROM (con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
        + vbCr + " WHERE  ((con_diario.fchasi) Is Null Or (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "')) " & nSQLIdLibro & nSQLTipoPersona & nSQLAjuste _
        + vbCr + " GROUP BY con_planctas.id, con_planctas.cuenta, con_planctas.descripcion " _
        + vbCr + " ORDER BY con_planctas.cuenta " _
        + vbCr + " ) AS SaldosIni "
    
    '--filtro para el where
    nSQLIdLibro = nSQLIdLibro & nSQLTipoPersona & nSQLAjuste & " AND (  (((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')))  OR  (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "')  OR  (con_diario.fchasi) is null  )"


    nSQL = nSQL _
        + vbCr + " ON con_planctas.id = SaldosIni.IdCta " _
        + vbCr + " WHERE con_planctas.id In (SELECT con_diario.idcue FROM con_diario " & IIf(nSQLIdLibro <> "", "WHERE " & Mid(nSQLIdLibro, 5), "") & "   ) " _
        + vbCr + " ORDER BY con_planctas.cuenta; "

    '--si seleccionar por periodo
    If opt_fecha(1).Value = True Then
        '--movimiento del periodo
        nSQL = Replace(nSQL, "(((con_diario.fchasi) Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "')))", "( con_diario.idmes>=" & mMesIni & " And con_diario.idmes <= " & mMesFin & " )")
        '--saldos iniciales
        nSQL = Replace(nSQL, "((con_diario.fchasi) Is Null Or (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "'))", "con_diario.idmes < " & mMesIni)
        '--filtro para where
        nSQL = Replace(nSQL, "OR  (con_diario.fchasi)<CDate('" & TxtFchIni.Valor & "')  OR  (con_diario.fchasi) is null ", " OR con_diario.idmes <= " & mMesIni)
        
    End If
    
    
    
    Me.MousePointer = vbHourglass
    DoEvents
    RST_Busq Rst, nSQL, xCon
    Dim mCol&
    If Rst.RecordCount <> 0 Then ProgressBar1.Max = Rst.RecordCount
    '--proceder a cargar los datos
    Do While Not Rst.EOF
        DoEvents
        Fg1(Indice).Rows = Fg1(Indice).Rows + 1
        '-----------------------------------------------
        ProgressBar1.Value = ProgressBar1.Value + 1
        DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then GoTo SALIR
        DoEvents
        For mCol = 0 To Rst.Fields.Count - 1
            
            If mCol <= 1 Then
                Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, mCol + 1) = NulosC(Rst(mCol))
            ElseIf mCol < 18 Then
                '--importes
                Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, mCol + 1) = Format(Rst(mCol), FORMAT_MONTO)
            Else
                '--iddest1,iddest2,idcta
                Fg1(Indice).TextMatrix(Fg1(Indice).Rows - 1, mCol + 1) = NulosN(Rst(mCol))
            
            
            End If
        Next mCol
        Rst.MoveNext
    Loop
SALIR:
    Frame11.Visible = False
    Me.MousePointer = vbDefault
    
End Sub






Private Sub CmdExpPDT_Click()
    ExportarPDT
End Sub



Sub ExportarPDT()
    Dim NomArch, xCad As String
    Dim A As Integer
    
    
    If Fg1(1).Rows = Fg1(1).FixedRows Then
        MsgBox "No hay registros para exportar", vbInformation, xTitulo
        Exit Sub
    End If
    
    
'***********
'''.Cell(flexcpText, 0, 3, 0, 4) = "Saldos Iniciales"
'''.Cell(flexcpText, 0, 5, 0, 6) = "Movimiento del Periodo"
'''.Cell(flexcpText, 0, 7, 0, 8) = "Sumas del Mayor"
'''.Cell(flexcpText, 0, 9, 0, 10) = "Saldos Al"
'''.Cell(flexcpText, 0, 11, 0, 12) = "Cuentas del Balance"
'''.Cell(flexcpText, 0, 13, 0, 14) = "Transferencias y Canc."
'''.Cell(flexcpText, 0, 15, 0, 16) = "Resultados x Naturaleza"
'''.Cell(flexcpText, 0, 17, 0, 18) = "Resultados x Función"

'***********
    
    
    
    If Fg1(1).Rows > Fg1(1).FixedRows Then
        '--cargar numero de ruc si esta vacio la variable
        If NumRUC = "" Then CargaDatosEmpresa
        '--indicar el nombre del archivo
        NomArch = "0664" & NumRUC & AnoTra & ".txt"
        
        Open Trim(App.Path) + "\" + NomArch For Output As #1
    
        For A = Fg1(1).FixedRows To Fg1(1).Rows - 4
            xCad = ""
            xCad = xCad & Fg1(1).TextMatrix(A, 1) + "|"
            xCad = xCad & NulosN(Fg1(1).TextMatrix(A, 3)) & "|" '--saldos iniciales - debe
            xCad = xCad & NulosN(Fg1(1).TextMatrix(A, 4)) & "|" '--saldos iniciales - haber
            xCad = xCad & NulosN(Fg1(1).TextMatrix(A, 5)) & "|" '--movimientos del ejercicio - debe
            xCad = xCad & NulosN(Fg1(1).TextMatrix(A, 6)) & "|" '--movimientos del ejercicio - haber
            xCad = xCad & NulosN(Fg1(1).TextMatrix(A, 13)) & "|" '--transferencias y Cancelaciones - debe
            xCad = xCad & NulosN(Fg1(1).TextMatrix(A, 14)) & "|" '--transferencias y Cancelaciones - haber
            
           
            
            Print #1, Trim(xCad)
        Next A
    End If
    
    Close #1
    
    MsgBox "El archivo para el DAOT se generó con éxito" & vbCr & "El archivo se grabó en el sgte. directorio: " & vbCr & Trim(App.Path) + "\" + NomArch, vbInformation, xTitulo
    
End Sub

Private Sub FraDetalle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    OrigFX = x
    OrigFY = y
    FraDetalle.ZOrder 0
End Sub

Private Sub FraDetalle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 0 Then
        With FraDetalle
            .Move .Left + x - OrigFX, .Top + y - OrigFY
        End With
    End If
End Sub

