VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmCanjesFacturas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caja y Bancos - Canje de Documentos"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   0
      TabIndex        =   12
      Top             =   390
      Width           =   11880
      _cx             =   20955
      _cy             =   12726
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
         Height          =   6795
         Left            =   12525
         TabIndex        =   17
         Top             =   375
         Width           =   11790
         Begin VB.CommandButton CmdBusProv 
            Height          =   230
            Left            =   2775
            Picture         =   "FrmCanjesFacturas.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   1230
            Width           =   210
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1395
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   2
            Text            =   "TxtNumSer"
            Top             =   870
            Width           =   915
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2670
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   3
            Text            =   "TxtNumDoc"
            Top             =   870
            Width           =   1440
         End
         Begin VB.Frame Frame4 
            Caption         =   "( Periodo )"
            Height          =   720
            Left            =   9480
            TabIndex        =   39
            Top             =   225
            Width           =   2010
            Begin VB.Label lblperiodo 
               Alignment       =   2  'Center
               Caption         =   "lblperiodo(1)"
               BeginProperty Font 
                  Name            =   "Courier"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   40
               Top             =   330
               Width           =   1740
            End
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   230
            Left            =   4725
            Picture         =   "FrmCanjesFacturas.frx":0132
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   570
            Width           =   210
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   4140
            Locked          =   -1  'True
            MaxLength       =   1
            TabIndex        =   1
            Text            =   "TxtIdMon"
            Top             =   540
            Width           =   825
         End
         Begin VB.TextBox TxtTotal4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   5010
            Locked          =   -1  'True
            TabIndex        =   34
            Text            =   "TxtTotal4"
            Top             =   3375
            Width           =   1095
         End
         Begin VB.TextBox TxtTotal3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   7965
            Locked          =   -1  'True
            TabIndex        =   33
            Text            =   "TxtTotal3"
            Top             =   6450
            Width           =   1095
         End
         Begin VB.CommandButton CmdBusCli 
            Height          =   230
            Left            =   2775
            Picture         =   "FrmCanjesFacturas.frx":0264
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   4140
            Width           =   210
         End
         Begin VB.TextBox TxtTotal2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   5010
            Locked          =   -1  'True
            TabIndex        =   26
            Text            =   "TxtTotal2"
            Top             =   6450
            Width           =   1095
         End
         Begin VB.TextBox TxtTotal1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   24
            Text            =   "TxtTotal1"
            Top             =   3375
            Width           =   1095
         End
         Begin VB.Frame Frame3 
            Height          =   1635
            Left            =   10095
            TabIndex        =   18
            Top             =   1695
            Width           =   1680
            Begin VB.CommandButton CmdAdd 
               Caption         =   "Agregar Documentos"
               Enabled         =   0   'False
               Height          =   495
               Left            =   135
               TabIndex        =   5
               Top             =   315
               Width           =   1350
            End
            Begin VB.CommandButton CmdDel 
               Caption         =   "Eliminar Documento"
               Enabled         =   0   'False
               Height          =   495
               Left            =   135
               TabIndex        =   19
               Top             =   870
               Width           =   1350
            End
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmi 
            Height          =   300
            Left            =   1395
            TabIndex        =   0
            Top             =   540
            Width           =   1260
            _ExtentX        =   2223
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
            Locked          =   -1  'True
            Valor           =   "18/07/2008"
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   1575
            Left            =   60
            TabIndex        =   6
            Top             =   1785
            Width           =   9915
            _cx             =   17489
            _cy             =   2778
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
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmCanjesFacturas.frx":0396
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
         Begin VSFlex7Ctl.VSFlexGrid Fg2 
            Height          =   1725
            Left            =   60
            TabIndex        =   11
            Top             =   4665
            Width           =   11700
            _cx             =   20637
            _cy             =   3043
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
            Cols            =   15
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmCanjesFacturas.frx":04E4
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
         Begin VSFlex7Ctl.VSFlexGrid VSFlexGrid1 
            Height          =   270
            Left            =   60
            TabIndex        =   32
            Top             =   4425
            Width           =   11700
            _cx             =   20637
            _cy             =   476
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
            FormatString    =   $"FrmCanjesFacturas.frx":06BD
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
         Begin VB.Frame Frame7 
            Height          =   675
            Left            =   7365
            TabIndex        =   28
            Top             =   3735
            Width           =   4410
            Begin VB.CommandButton CmdProcesar 
               Caption         =   "Canjear Documentos"
               Enabled         =   0   'False
               Height          =   435
               Left            =   2925
               TabIndex        =   10
               Top             =   165
               Width           =   1350
            End
            Begin VB.CommandButton CmdDelDocEmi 
               Caption         =   "Eliminar Documento"
               Enabled         =   0   'False
               Height          =   435
               Left            =   1470
               TabIndex        =   9
               Top             =   165
               Width           =   1350
            End
            Begin VB.CommandButton CmdAddDocEmi 
               Caption         =   "Agregar Documentos"
               Enabled         =   0   'False
               Height          =   435
               Left            =   75
               TabIndex        =   8
               Top             =   165
               Width           =   1350
            End
         End
         Begin VB.TextBox TxtRucPro 
            Height          =   300
            Left            =   1395
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   4
            Text            =   "TxtRucPro"
            Top             =   1200
            Width           =   1620
         End
         Begin VB.TextBox TxtRucCli 
            Height          =   300
            Left            =   1395
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   7
            Text            =   "TxtRucCli"
            Top             =   4110
            Width           =   1620
         End
         Begin VB.Label LblCliente 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblCliente"
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
            Left            =   3015
            TabIndex        =   48
            Top             =   4110
            Width           =   4080
         End
         Begin VB.Label LblProveedor 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblProveedor"
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
            Left            =   3015
            TabIndex        =   47
            Top             =   1200
            Width           =   3960
         End
         Begin VB.Label LblTitulo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor"
            Height          =   195
            Left            =   75
            TabIndex        =   46
            Top             =   1275
            Width           =   735
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            Height          =   90
            Left            =   2430
            Top             =   975
            Width           =   105
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Documento"
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   44
            Top             =   960
            Width           =   1275
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00800000&
            BorderWidth     =   2
            X1              =   60
            X2              =   11700
            Y1              =   3720
            Y2              =   3720
         End
         Begin VB.Label LblTipCam2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Cambio"
            Height          =   195
            Left            =   7170
            TabIndex        =   43
            Top             =   1260
            Width           =   1110
         End
         Begin VB.Label LblTipoCambio 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipoCambio"
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
            Left            =   8400
            TabIndex        =   42
            Top             =   1170
            Width           =   1980
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   4
            Left            =   3450
            TabIndex        =   37
            Top             =   615
            Width           =   585
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
            Height          =   300
            Left            =   4950
            TabIndex        =   36
            Top             =   540
            Width           =   1155
         End
         Begin VB.Label LblIdCliente 
            AutoSize        =   -1  'True
            Caption         =   "LblIdCliente"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   765
            TabIndex        =   31
            Top             =   3915
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   30
            Top             =   4215
            Width           =   480
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Total ==>"
            Height          =   195
            Index           =   1
            Left            =   4215
            TabIndex        =   27
            Top             =   6495
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Total ==>"
            Height          =   195
            Left            =   4215
            TabIndex        =   25
            Top             =   3465
            Width           =   675
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Canje de Documentos"
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
            Left            =   60
            TabIndex        =   23
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emisión"
            Height          =   195
            Index           =   3
            Left            =   75
            TabIndex        =   22
            Top             =   615
            Width           =   1260
         End
         Begin VB.Label LblIdProveedor 
            AutoSize        =   -1  'True
            Caption         =   "LblIdProveedor"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   5310
            TabIndex        =   21
            Top             =   990
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label LblTituloDoc 
            AutoSize        =   -1  'True
            Caption         =   "Documentos del Proveedor"
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
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   20
            Top             =   1560
            Width           =   2310
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6795
         Left            =   45
         TabIndex        =   13
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6360
            Left            =   30
            TabIndex        =   14
            Top             =   345
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11218
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Num.Reg."
            Columns(0).DataField=   "numreg"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Número Doc."
            Columns(1).DataField=   "numerodoc"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Emi."
            Columns(2).DataField=   "fchemi"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Cliente"
            Columns(3).DataField=   "nomcli"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Proveedor"
            Columns(4).DataField=   "nompro"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "M"
            Columns(5).DataField=   "monabrev"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Imp. Canjeado"
            Columns(6).DataField=   "impcan"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1720"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1640"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2434"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2355"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1773"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1693"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=5424"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=5345"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=4921"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=4842"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=847"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=767"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=2381"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2302"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=514"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
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
            HeadLines       =   1.5
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
            _StyleDefs(11)  =   ":id=2,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
            _StyleDefs(64)  =   "Named:id=33:Normal"
            _StyleDefs(65)  =   ":id=33,.parent=0"
            _StyleDefs(66)  =   "Named:id=34:Heading"
            _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(68)  =   ":id=34,.wraptext=-1"
            _StyleDefs(69)  =   "Named:id=35:Footing"
            _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(71)  =   "Named:id=36:Selected"
            _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=37:Caption"
            _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(75)  =   "Named:id=38:HighlightRow"
            _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=39:EvenRow"
            _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(79)  =   "Named:id=40:OddRow"
            _StyleDefs(80)  =   ":id=40,.parent=33"
            _StyleDefs(81)  =   "Named:id=41:RecordSelector"
            _StyleDefs(82)  =   ":id=41,.parent=34"
            _StyleDefs(83)  =   "Named:id=42:FilterBar"
            _StyleDefs(84)  =   ":id=42,.parent=33"
         End
         Begin VB.Label lblperiodo 
            Caption         =   "lblperiodo(0)"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   0
            Left            =   9450
            TabIndex        =   41
            Top             =   75
            Width           =   1980
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Canje de Documentos"
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
            TabIndex        =   16
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label LblMes 
            Caption         =   "LblMes"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   300
            Left            =   8835
            TabIndex        =   15
            Top             =   30
            Width           =   1275
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Listado"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Registro"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7410
         Top             =   45
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483637
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjesFacturas.frx":07AB
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjesFacturas.frx":0CEF
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjesFacturas.frx":1081
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjesFacturas.frx":1205
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjesFacturas.frx":1659
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjesFacturas.frx":1771
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjesFacturas.frx":1CB5
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjesFacturas.frx":21F9
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjesFacturas.frx":230D
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjesFacturas.frx":2421
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjesFacturas.frx":2875
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCanjesFacturas.frx":29E1
               Key             =   "IMG11"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmCanjesFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstFrm As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim QueHace As Integer
Dim Agregando As Boolean
Dim xHorIni As Date

Dim mMesActivo As Integer '--indica el mes activo
Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim fCierrePeriodo As Boolean '--indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO


Private Sub CmdAdd_Click()
    On Error GoTo error
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "No ha especificado la Moneda", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMon.SetFocus
        Exit Sub
    End If
    If NulosN(LblIdProveedor.Caption) = 0 Then
        MsgBox "No ha especificado el proveedor", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtRucPro.SetFocus
        Exit Sub
    End If

    Dim xCampos(6, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim nSQLId As String
    Dim nSQL As String
    
    xCampos(0, 0) = "Tipo Doc.":       xCampos(0, 1) = "abrev":     xCampos(0, 2) = "1000":    xCampos(0, 3) = "C":     xCampos(0, 4) = "N"
    xCampos(1, 0) = "Nº Documento":    xCampos(1, 1) = "numdoc":    xCampos(1, 2) = "2000":    xCampos(1, 3) = "C":     xCampos(1, 4) = "S"
    xCampos(2, 0) = "Fch.Emision":     xCampos(2, 1) = "fchdoc":    xCampos(2, 2) = "1200":    xCampos(2, 3) = "C":     xCampos(2, 4) = "N"
    xCampos(3, 0) = "M":               xCampos(3, 1) = "simbolo":   xCampos(3, 2) = "450":     xCampos(3, 3) = "C":     xCampos(3, 4) = "N"
    xCampos(4, 0) = "Importe":         xCampos(4, 1) = "imptotdoc": xCampos(4, 2) = "1100":    xCampos(4, 3) = "N":     xCampos(4, 4) = "N"
    xCampos(5, 0) = "Saldo":           xCampos(5, 1) = "impsal":    xCampos(5, 2) = "1100":    xCampos(5, 3) = "N":     xCampos(5, 4) = "N"
    
    '--generar el script para no
    nSQLId = GRID_GENERAR_SQL_ID(Fg1, 8, "com_compras.id", " NOT IN ")
    If nSQLId <> "" Then nSQLId = " AND " + nSQLId
    '--------
    
    nSQL = "SELECT mae_documento.abrev, com_compras!numser+'-'+com_compras!numdoc AS numdoc, com_compras.impsal, com_compras.fchdoc, " _
        & " com_compras.fchven, com_compras.idpro, com_compras.imptot AS imptotdoc, com_compras.id, con_planctas.id as idcue, con_planctas.cuenta, " _
        & " mae_moneda.simbolo,com_compras.idmon " _
        & " FROM (mae_documento RIGHT JOIN (mae_moneda RIGHT JOIN com_compras ON mae_moneda.id = com_compras.idmon) ON mae_documento.id = com_compras.tipdoc) INNER JOIN (con_planctas INNER JOIN mae_documentocta  " _
        & " ON con_planctas.id = mae_documentocta.idcuen) ON (com_compras.tipdoc = mae_documentocta.iddoc) AND (com_compras.idmon = mae_documentocta.idmon) " _
        & " WHERE com_compras.idmon = " & NulosN(TxtIdMon.Text) & "  and (((com_compras.impsal) > 0) And ((com_compras.idpro) = " & NulosN(LblIdProveedor.Caption) & ") And ((con_planctas.cuenta) Like '" & 42 & "%')) " _
        & nSQLId _
        & " ORDER BY com_compras.numser+'-'+com_compras.numdoc ; "
   
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando documentos de venta del Cliente", "numdoc", "numdoc", CualquierParte
    
    If xRs.State = 0 Then GoTo Salir
    If xRs.RecordCount = 0 Then GoTo Salir
    Fg1.Rows = 1
        
    Fg1.Rows = Fg1.Rows + 1
    Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(xRs("abrev"))
    Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(xRs("numdoc"))
    Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(xRs("fchdoc"))
    Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(xRs("fchven"))
    Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(xRs("simbolo")) '----
    Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(xRs("imptotdoc")), FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(xRs("impsal")), FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosN(xRs("id"))
    Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosN(xRs("idcue"))
    Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(NulosN(xRs("impsal")), FORMAT_MONTO)

    HallarTotales
Salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "CmdAdd_Click"
End Sub

Private Sub HallarTotales()
    '--DOCUMENTOS DEL PROVEEDOR
    TxtTotal4.Text = Format(GRID_SUMAR_COL(Fg1, 6), FORMAT_MONTO)
    TxtTotal1.Text = Format(GRID_SUMAR_COL(Fg1, 7), FORMAT_MONTO)
    '--DOCUMENTOS DEL CLIENTE
    TxtTotal2.Text = Format(GRID_SUMAR_COL(Fg2, 6), FORMAT_MONTO)
    TxtTotal3.Text = Format(GRID_SUMAR_COL(Fg2, 8), FORMAT_MONTO)
End Sub

Private Sub CmdAddDocEmi_Click()
    On Error GoTo error
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "No ha especificado la Moneda", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMon.SetFocus
        Exit Sub
    End If
    If NulosN(LblIdCliente.Caption) = 0 Then
        MsgBox "No ha especificado el nombre del Cliente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtRucCli.SetFocus
        Exit Sub
    End If

    Dim xCampos(6, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim nSQLId  As String
    Dim nSQL As String
    
    xCampos(0, 0) = "Tipo Doc.":       xCampos(0, 1) = "abrev":     xCampos(0, 2) = "1000":    xCampos(0, 3) = "C":     xCampos(0, 4) = "N"
    xCampos(1, 0) = "Nº Documento":    xCampos(1, 1) = "numdoc":    xCampos(1, 2) = "2000":    xCampos(1, 3) = "C":     xCampos(1, 4) = "S"
    xCampos(2, 0) = "Fch.Emision":     xCampos(2, 1) = "fchdoc":    xCampos(2, 2) = "1200":    xCampos(2, 3) = "C":     xCampos(2, 4) = "N"
    xCampos(3, 0) = "M":               xCampos(3, 1) = "simbolo":   xCampos(3, 2) = "450":     xCampos(3, 3) = "C":     xCampos(3, 4) = "N"
    xCampos(4, 0) = "Importe":         xCampos(4, 1) = "imptotdoc": xCampos(4, 2) = "1100":    xCampos(4, 3) = "N":     xCampos(4, 4) = "N"
    xCampos(5, 0) = "Saldo":           xCampos(5, 1) = "impsal":    xCampos(5, 2) = "1100":    xCampos(5, 3) = "N":     xCampos(5, 4) = "N"

    '--generar el script para no
    nSQLId = GRID_GENERAR_SQL_ID(Fg2, 11, "vta_ventas.id", " NOT IN ")
    If nSQLId <> "" Then nSQLId = " AND " + nSQLId
    '--------
    
    nSQL = "SELECT [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, mae_documento.abrev, vta_ventas.fchdoc, vta_ventas.fchven, " _
        & " vta_ventas.imptotdoc, vta_ventas.impsal, vta_ventas.id,con_planctas.id as idcue, con_planctas.cuenta, mae_moneda.simbolo " _
        & " FROM ((mae_documento RIGHT JOIN vta_ventas ON mae_documento.id = vta_ventas.tipdoc) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) INNER JOIN (con_planctas INNER JOIN mae_documentocta  " _
        & " ON con_planctas.id = mae_documentocta.idcuen) ON (vta_ventas.tipdoc = mae_documentocta.iddoc) AND (vta_ventas.idmon = mae_documentocta.idmon) " _
        & " WHERE vta_ventas.idmon = " & NulosN(TxtIdMon.Text) & "  and  (((vta_ventas.impsal) <> 0) And ((vta_ventas.idcli) = " & NulosN(LblIdCliente.Caption) & ") And mae_documentocta.tipope=-1 And " _
        & " ((con_planctas.cuenta) Like '" & 12 & "%')) " & nSQLId & " ORDER BY vta_ventas.numser+'-'+vta_ventas.numdoc"
    DoEvents
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Documentos Emitidos", "numdoc", "numdoc", CualquierParte
    
    If xRs.State = 0 Then GoTo Salir
    If xRs.RecordCount = 0 Then GoTo Salir
    Fg2.Rows = 1
    
    Fg2.Rows = Fg2.Rows + 1
    Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(xRs("abrev"))
    Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosC(xRs("numdoc"))
    Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(xRs("fchdoc"), "dd/mm/yy")
    Fg2.TextMatrix(Fg2.Rows - 1, 4) = NulosC(xRs("simbolo"))
    Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(NulosN(xRs("imptotdoc")), FORMAT_MONTO)
    Fg2.TextMatrix(Fg2.Rows - 1, 6) = Format(NulosN(xRs("impsal")), FORMAT_MONTO)
    Fg2.TextMatrix(Fg2.Rows - 1, 10) = Format(NulosN(xRs("impsal")), FORMAT_MONTO)
    Fg2.TextMatrix(Fg2.Rows - 1, 11) = NulosN(xRs("id"))
    Fg2.TextMatrix(Fg2.Rows - 1, 13) = NulosN(xRs("idcue"))
             
    If Fg2.Rows > 1 Then
        With Fg2
            .Select 1, 1, Fg2.Rows - 1, 5
            .FillStyle = flexFillRepeat
            .CellBackColor = &HDBF8F9
        
            .Select 1, 9, Fg2.Rows - 1, 9
            .FillStyle = flexFillRepeat
            .CellBackColor = &HDBF8F9
            
            .Select 1, 1, 1, 1
        End With
    End If
Salir:
    HallarTotales
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "CmdAddDocEmi_Click"
End Sub

Private Sub CmdBusCli_Click()
    If QueHace = 3 Then Exit Sub
    On Error GoTo error
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    Dim nSQL As String
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre":      xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":   xCampos(1, 1) = "numruc":        xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"
    '--MOSTRAR CLIENTES CON SALDO <> 0
    nSQL = "SELECT DISTINCT mae_cliente.numruc, mae_cliente.id, mae_cliente.nombre, mae_cliente.ageret " _
        + vbCr + " FROM mae_cliente INNER JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli " _
        + vbCr + " WHERE (((vta_ventas.impsal)<>0)) " _
        + vbCr + " ORDER BY mae_cliente.nombre ASC ;"
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Clientes", "nombre", "nombre", Principio
    If xRs.State = 0 Then GoTo Salir
    If xRs.RecordCount = 0 Then GoTo Salir
    
    LblIdCliente.Tag = LblIdCliente.Caption
    TxtRucCli.Text = NulosC(xRs("numruc"))
    LblCliente.Caption = NulosC(xRs("nombre"))
    LblIdCliente.Caption = NulosN(xRs("id"))
    If LblIdCliente.Tag <> LblIdCliente.Caption Then
        Fg2.Rows = 1
        HallarTotales
    End If
    
    'CargarFacturasCliente
Salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "CmdBusCli_Click"
End Sub

Private Sub CmdBusMon_Click()
    If QueHace = 3 Then Exit Sub
    On Error GoTo error
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    Dim nSQL As String
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre":      xCampos(0, 1) = "descripcion":     xCampos(0, 2) = "4500":      xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":   xCampos(1, 1) = "id":              xCampos(1, 2) = "1400":      xCampos(1, 3) = "C"

    nSQL = "SELECT * FROM mae_moneda ORDER BY descripcion"
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Moneda", "descripcion", "descripcion", Principio
    
    If xRs.State = 0 Then GoTo Salir:
    If xRs.RecordCount = 0 Then GoTo Salir:
    TxtIdMon.Tag = TxtIdMon.Text
    TxtIdMon.Text = NulosN(xRs("id"))
    LblMoneda.Caption = NulosN(xRs("descripcion"))
    If TxtIdMon.Tag <> TxtIdMon.Text Then
        Fg1.Rows = 1
        Fg2.Rows = 1
        HallarTotales
    End If
    TxtNumSer.SetFocus
    
Salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "CmdBusMon_Click"
End Sub

Private Sub CmdBusProv_Click()
    If QueHace = 3 Then Exit Sub
    On Error GoTo error
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    Dim nSQL As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre":      xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":   xCampos(1, 1) = "numruc":        xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"

    '--MOSTRAR SOLO LOS PROVEEDORES QUE TENGAN SALDO
    nSQL = "SELECT DISTINCT mae_prov.numruc, mae_prov.id, mae_prov.nombre, mae_prov.ageret " _
        + vbCr + " FROM mae_prov INNER JOIN com_compras ON mae_prov.id = com_compras.idpro " _
        + vbCr + " WHERE (((com_compras.impsal) <> 0)) " _
        + vbCr + " ORDER BY mae_prov.nombre;"
        
            
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Clientes como Proveedor", "nombre", "nombre", Principio
    
    If xRs.State = 0 Then GoTo Salir
    If xRs.RecordCount = 0 Then GoTo Salir
    LblIdProveedor.Tag = LblIdProveedor.Caption
    TxtRucPro.Text = NulosC(xRs("numruc"))
    LblProveedor.Caption = NulosC(xRs("nombre"))
    LblIdProveedor.Caption = NulosN(xRs("id"))
    If LblIdProveedor.Tag <> LblIdProveedor.Caption Then
        Fg1.Rows = 1
        Fg2.Rows = 1
        HallarTotales
    End If
    'CargarFacturasCliente
    TxtRucCli.SetFocus
Salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "CmdBusProv_Click"
End Sub

Private Sub CmdDel_Click()
    If Fg1.Row <= 0 Then Exit Sub
    If Fg1.Rows = 1 Then
        MsgBox "No ha documentos para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Fg1.RemoveItem Fg1.Row
    Fg2.Rows = 1
    HallarTotales
End Sub

Private Sub CmdProcesar_Click()
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado los documentos de compra a canjear", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Sub
    End If
    
    If Fg2.Rows = 1 Then
        MsgBox "No ha especificado que documentos de venta a canjear", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg2.SetFocus
        Exit Sub
    End If
    
    Dim A, B As Integer
    Dim Saldo As Double
    
    '--RESTAURAR EL SALDO DE DOCUMENTOS DEL PROVEEDOR
    For A = 1 To Fg1.Rows - 1
        Fg1.TextMatrix(A, 7) = Fg1.TextMatrix(A, 10)
    Next A
    '------------
    
    For A = 1 To Fg2.Rows - 1
        For B = 1 To Fg1.Rows - 1
            If NulosN(Fg1.TextMatrix(B, 6)) > 0 Then 'si es saldo en el documento del proveedor es mayor a 0
                
                If NulosN(Fg1.TextMatrix(B, 7)) > NulosN(Fg2.TextMatrix(A, 6)) Then
                    Saldo = NulosN(Fg2.TextMatrix(A, 6)) - NulosN(Fg1.TextMatrix(B, 7))
                    If Saldo < 0 Then 'si el saldo es negativo, quiere decir que esta quedando saldo en el documento del proveedor que se esta canjeando
                        Fg2.TextMatrix(A, 9) = Format(Abs(Saldo), FORMAT_MONTO)
                        Fg1.TextMatrix(B, 7) = Format(Abs(Saldo), FORMAT_MONTO)
                        Fg2.TextMatrix(A, 8) = Fg2.TextMatrix(A, 6)
                    Else
                        'si el saldo es positivo quiere decir que el documento del proveedor se quedo sin saldo
                        Fg2.TextMatrix(A, 9) = "0.00"
                        Fg2.TextMatrix(A, 8) = Fg1.TextMatrix(B, 7)
                        Fg1.TextMatrix(B, 7) = "0.00"
                    End If
                    Fg2.TextMatrix(A, 7) = Fg1.TextMatrix(B, 2)
                    Fg2.TextMatrix(A, 10) = Format(NulosN(Fg2.TextMatrix(A, 6)) - NulosN(Fg2.TextMatrix(A, 8)), FORMAT_MONTO)
                    Fg2.TextMatrix(A, 12) = Fg1.TextMatrix(B, 8)
                    Fg2.TextMatrix(A, 14) = Fg1.TextMatrix(B, 9)
                    If Saldo < 0 Then
                        If (NulosN(Fg2.TextMatrix(A, 6)) - NulosN(Fg2.TextMatrix(A, 8))) = 0 Then
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Else
                    Saldo = NulosN(Fg2.TextMatrix(A, 6)) - NulosN(Fg1.TextMatrix(B, 7))
                    If Saldo < 0 Then 'si el saldo es negativo, quiere decir que esta quedando saldo en el documento del proveedor que se esta canjeando
                        Fg2.TextMatrix(A, 9) = Format(Abs(Saldo), FORMAT_MONTO)
                        Fg1.TextMatrix(B, 7) = Format(Abs(Saldo), FORMAT_MONTO)
                        Fg2.TextMatrix(A, 8) = Fg2.TextMatrix(A, 6)
                    Else
                        'si el saldo es positivo quiere decir que el documento del proveedor se quedo sin saldo
                        Fg2.TextMatrix(A, 9) = "0.00"
                        Fg2.TextMatrix(A, 8) = Fg1.TextMatrix(B, 7)
                        Fg1.TextMatrix(B, 7) = "0.00"
                    End If
                    Fg2.TextMatrix(A, 7) = Fg1.TextMatrix(B, 2)
                    Fg2.TextMatrix(A, 10) = Format(NulosN(Fg2.TextMatrix(A, 6)) - NulosN(Fg2.TextMatrix(A, 8)), FORMAT_MONTO)
                    Fg2.TextMatrix(A, 12) = Fg1.TextMatrix(B, 8)
                    Fg2.TextMatrix(A, 14) = Fg1.TextMatrix(B, 9)
                    
                    If Saldo < 0 Then
                        If (NulosN(Fg2.TextMatrix(A, 6)) - NulosN(Fg2.TextMatrix(A, 8))) = 0 Then
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                End If
            End If
        Next B
    Next A
    HallarTotales
End Sub

Private Sub CmdDelDocEmi_Click()
    If Fg2.Row <= 0 Then Exit Sub

    If Fg2.Rows = 1 Then
        MsgBox "No ha documentos para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Fg2.RemoveItem Fg2.Row
    HallarTotales
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstFrm("id")), xCon
    End If
End Sub

Private Sub Fg1_EnterCell()
    Fg1.Editable = flexEDNone
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 45 Then '--AGREGAR
        CmdAdd_Click
    End If
    If KeyCode = 46 Then '--ELIMINAR
        CmdDel_Click
    End If
    
End Sub

Private Sub Fg2_EnterCell()
    Fg2.Editable = flexEDNone
End Sub

Private Sub Fg2_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub Fg2_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 45 Then '--AGREGAR
        CmdAddDocEmi_Click
    End If
    If KeyCode = 46 Then '--ELIMINAR
        CmdDelDocEmi_Click
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
   
         SeEjecuto = True
         
         mMesActivo = xMes
         
         '--Almacenar temporalmente el codigo del menu
         IdMenuActivo = xIdMenu
        
         pCargarGrid
    
    End If

End Sub


Sub Blanquea()
    LblTipoCambio.Caption = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtFchEmi.Valor = ""
    TxtRucPro.Text = ""
    TxtRucCli.Text = ""
    TxtTotal1.Text = ""
    TxtTotal2.Text = ""
    TxtTotal3.Text = ""
    TxtTotal4.Text = ""
    TxtIdMon.Text = ""
    LblMoneda.Caption = ""
End Sub

Sub Bloquea()
    TxtFchEmi.Locked = Not TxtFchEmi.Locked
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    TxtRucPro.Locked = Not TxtRucPro.Locked
    TxtRucCli.Locked = Not TxtRucCli.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    '--DOC PROVEEDOR
    CmdAdd.Enabled = Not CmdAdd.Enabled
    CmdDel.Enabled = Not CmdDel.Enabled
    '--DOC CLIENTE
    CmdAddDocEmi.Enabled = Not CmdAddDocEmi.Enabled
    CmdDelDocEmi.Enabled = Not CmdDelDocEmi.Enabled
    CmdProcesar.Enabled = Not CmdProcesar.Enabled
    
End Sub

Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Nuevo()
    QueHace = 1
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Bloquea
    Blanquea
    Label5.Caption = "Agregando Canje de Documentos"
    Fg1.Rows = 1
    Fg2.Rows = 1
    xHorIni = Time
    TxtFchEmi.SetFocus
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    
    Dg1.Columns("fchemi").NumberFormat = FORMAT_DATE:
    Dg1.Columns("impcan").NumberFormat = FORMAT_MONTO:
    
    Fg1.ColWidth(8) = 0
    Fg1.ColWidth(9) = 0
    Fg1.ColWidth(10) = 0

    Fg2.ColWidth(11) = 0
    Fg2.ColWidth(12) = 0
    Fg2.ColWidth(13) = 0
    Fg2.ColWidth(14) = 0

    QueHace = 3
    
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    TabOne1.CurrTab = 0
    Fg1.SelectionMode = flexSelectionByRow
    Fg2.SelectionMode = flexSelectionByRow
End Sub

Sub CargarFacturasCliente()
    Dim rst As New ADODB.Recordset
    Dim nSQL As String
    
    nSQL = "SELECT  com_compras.id, com_compras.idpro, mae_documento.abrev, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, com_compras.fchdoc, " _
        & " com_compras.fchven, com_compras.imptot, com_compras.impsal FROM mae_documento RIGHT JOIN com_compras ON mae_documento.id = com_compras.tipdoc " _
        & " WHERE (((com_compras.idpro)=" & Val(LblIdProveedor.Caption) & ") AND ((com_compras.impsal)<>0))"
    
    RST_Busq rst, nSQL, xCon
    
    If rst.RecordCount <> 0 Then rst.MoveFirst
       
    Do While Not rst.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(rst("abrev"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(rst("numdoc"))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(rst("fchemi"))
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(rst("fchven"))
        Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(NulosN(rst("imptot")), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(rst("impsal")), FORMAT_MONTO)
        rst.MoveNext
    Loop
    Set rst = Nothing
End Sub

Sub Modificar()
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay registro para modificar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    QueHace = 2
    Label5.Caption = "Modificando Canje de Documento"
    Bloquea
    ActivaTool
    
    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    Else
        If CmdAdd.Enabled = False Then Bloquea
    End If
    
    TabOne1.TabEnabled(0) = False
    
    Fg1.Editable = flexEDKbdMouse
    Fg1.SelectionMode = flexSelectionFree
    Fg2.Editable = flexEDKbdMouse
    Fg2.SelectionMode = flexSelectionFree
    
    Agregando = False
    
    xHorIni = Time

    TxtFchEmi.SetFocus

End Sub

Sub Cancelar()
    ActivaTool
    Label5.Caption = "Detalle del Canje de Documentos"
    QueHace = 3
    Bloquea
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
End Sub

Function Grabar() As Boolean
    
    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "Grabar", "Modificar") + " el canje del Documento", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
''    Dim RstDia As New ADODB.Recordset
    Dim xId As Double
    Dim A&
    Dim xNumAsiento As String
    Dim nSQL As String
On Error GoTo LaCague

    xCon.BeginTrans
    

    If QueHace = 1 Then
        xId = HallaCodigoTabla("con_canjes", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM con_canjes", xCon
       
''''        xNumAsiento = NuevoNumAsiento(8, mMesActivo, xCon)
        RstCab.AddNew
    Else
        xId = RstFrm("id")
        
''''        xNumAsiento = DevuelveNumAsiento(8, RstFrm("id"), mMesActivo, xCon)
''''
''''        If xNumAsiento = "" Then xNumAsiento = NuevoNumAsiento(8, mMesActivo, xCon)
        '*********************************************************************************************
        Dim RstTmp As New ADODB.Recordset

        nSQL = "SELECT con_canjesdet.tipo, con_canjesdet.iddoc, con_canjesdet.iddoccan, con_canjesdet.impcan " _
                + vbCr + " FROM con_canjesdet " _
                + vbCr + " WHERE (((con_canjesdet.tipo)=1) AND ((con_canjesdet.idcan)=" & xId & "))"
                
        RST_Busq RstTmp, nSQL, xCon
        If RstTmp.RecordCount <> 0 Then
            RstTmp.MoveFirst
            Do While Not RstTmp.EOF
                'actualizamos el saldo del documento de venta y del documento de compra
                xCon.Execute "UPDATE com_compras SET com_compras.impsal = [com_compras]![impsal] + " & NulosN(RstTmp.Fields("impcan")) & " WHERE (com_compras.id = " & RstTmp.Fields("iddoccan") & " )"
                'actualizamos el saldo del documento de venta y del documento de compra
                xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = [vta_ventas]![impsal] + " & NulosN(RstTmp.Fields("impcan")) & " WHERE (vta_ventas.id =" & RstTmp.Fields("iddoc") & ")"
                '
                RstTmp.MoveNext
            Loop
        End If
        Set RstTmp = Nothing
        
        xCon.Execute "DELETE * FROM con_canjesdet WHERE idcan = " & xId & ""
        'eliminamos los asientos contables
''''        xCon.Execute "DELETE * FROM con_diario WHERE ((idmes = " & mMesActivo & ") and (idlib = 8) AND (idmov = " & xId & ")) ;"
        '*********************************************************************************************
        RST_Busq RstCab, "SELECT * FROM con_canjes WHERE id = " & xId & "", xCon


    End If
    '**************
    RST_Busq RstDet, "SELECT TOP 1 * FROM con_canjesdet", xCon
''''    RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
    '**************
    
    RstCab("ano") = AnoTra
    RstCab("idmes") = mMesActivo
    RstCab("idlib") = 8
''''    RstCab("numreg") = Format(mMesActivo, "00") + xNumAsiento
    If mMesActivo <> 0 And mMesActivo <> 13 Then
        RstCab("fchreg") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
    End If
    
    RstCab("id") = xId
    RstCab("fchemi") = CDate(TxtFchEmi.Valor)
    RstCab("idpro") = NulosN(LblIdProveedor.Caption)
    RstCab("idcli") = NulosN(LblIdCliente.Caption)
    RstCab("impcan") = NulosN(TxtTotal3.Text)
    
    RstCab("numser") = NulosC(TxtNumSer.Text)
    RstCab("numdoc") = NulosC(TxtNumDoc.Text)
    
    RstCab("idmon") = NulosN(TxtIdMon.Text)
    
    RstCab.Update
    
    'grabamos el detalle del canje
    'grabamos los documentos de compra a canjear
    For A = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("idcan") = xId
        RstDet("tipo") = 2
        RstDet("iddoc") = NulosN(Fg1.TextMatrix(A, 8))
        RstDet("impdoc") = NulosN(Fg1.TextMatrix(A, 6))
        RstDet("saldo") = NulosN(Fg1.TextMatrix(A, 7))
        
        RstDet.Update
    Next A
    
    'grabamos los documentos de venta canjeados
    For A = 1 To Fg2.Rows - 1
        RstDet.AddNew
        RstDet("idcan") = xId
        RstDet("tipo") = 1
        RstDet("iddoc") = NulosN(Fg2.TextMatrix(A, 11))
        RstDet("impdoc") = NulosN(Fg2.TextMatrix(A, 5))
        RstDet("saldo") = NulosN(Fg2.TextMatrix(A, 6))
        RstDet("iddoccan") = NulosN(Fg2.TextMatrix(A, 12))
        RstDet("impcan") = NulosN(Fg2.TextMatrix(A, 8))
        RstDet("impsalcan") = NulosN(Fg2.TextMatrix(A, 9))
        RstDet.Update
        
        'grabamos el diario
        
''''        'GRABAMOS LA CUENTA DEBE
''''        RstDia.AddNew
''''        RstDia("año") = AnoTra
''''        RstDia("idmes") = mMesActivo
''''        RstDia("idlib") = 8    'libro de canjes de facturas
''''        RstDia("idmov") = xId
''''        RstDia("numasi") = xNumAsiento
''''        RstDia("tc") = NulosN(LblTipoCambio.Caption)
''''        If mMesActivo = 13 Then
''''            RstDia("fchasi") = CDate("31/12/" + AnoTra)
''''        Else
''''            RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
''''        End If
''''        RstDia("idcue") = NulosN(Fg2.TextMatrix(A, 13))
''''        RstDia("iddocpro") = NulosN(Fg2.TextMatrix(A, 11))
''''        RstDia("correlativo") = A
''''
''''        If NulosN(TxtIdMon.Text) = 1 Then '--SOLES
''''            RstDia("impdebsol") = NulosN(Fg2.TextMatrix(A, 8))
''''            RstDia("impdebdol") = 0
''''        Else
''''            RstDia("impdebsol") = NulosN(Fg2.TextMatrix(A, 8)) * NulosN(LblTipoCambio.Caption)
''''            RstDia("impdebdol") = NulosN(Fg2.TextMatrix(A, 8))
''''        End If
''''        RstDia("fchdoc") = CDate(TxtFchEmi.Valor)
''''
''''        RstDia.Update
''''
''''        'GRABAMOS LA CUENTA HABER
''''        RstDia.AddNew
''''        RstDia("año") = AnoTra
''''        RstDia("idmes") = mMesActivo
''''        RstDia("idlib") = 8    'libro de canje de facturas
''''        RstDia("idmov") = xId
''''        RstDia("numasi") = xNumAsiento
''''        RstDia("tc") = NulosN(LblTipoCambio.Caption)
''''        If mMesActivo = 13 Then
''''            RstDia("fchasi") = CDate("31/12/" + AnoTra)
''''        Else
''''            RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
''''        End If
''''        RstDia("idcue") = NulosN(Fg2.TextMatrix(A, 14))
''''        RstDia("iddocpro") = NulosN(Fg2.TextMatrix(A, 12))
''''        RstDia("correlativo") = A
''''
''''        If NulosN(TxtIdMon.Text) = 1 Then '--SOLES
''''            RstDia("imphabsol") = NulosN(Fg2.TextMatrix(A, 8))
''''            RstDia("imphabdol") = 0
''''        Else
''''            RstDia("imphabsol") = NulosN(Fg2.TextMatrix(A, 8)) * NulosN(LblTipoCambio.Caption)
''''            RstDia("imphabdol") = NulosN(Fg2.TextMatrix(A, 8))
''''        End If
''''        RstDia("fchdoc") = CDate(TxtFchEmi.Valor)
''''        RstDia.Update
        
        'actualizamos el saldo del documento de compra
        xCon.Execute "UPDATE com_compras SET com_compras.impsal = [com_compras]![impsal] - " & NulosN(Fg2.TextMatrix(A, 8)) & " WHERE (com_compras.id = " & NulosN(Fg2.TextMatrix(A, 12)) & ") ;"
        
        'actualizamos el saldo del documento de venta
        xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = [vta_ventas]![impsal] - " & NulosN(Fg2.TextMatrix(A, 8)) & " WHERE (vta_ventas.id =" & NulosN(Fg2.TextMatrix(A, 11)) & ") ;"

    Next A
    
    '--generamos es asiento
    xNumAsiento = GenerarAsiento(xCon, 8, CDbl(xId), AnoTra, mMesActivo, 1, 2)
    If xNumAsiento = "" Then GoTo LaCague
    '---------------------------------------------------------------------------
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId

    Me.MousePointer = vbDefault
    
    xCon.CommitTrans
    
    Set RstCab = Nothing
    Set RstDet = Nothing
''''    Set RstDia = Nothing
    
''''    MsgBox "El canje del Documento se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito" + vbCr _
''''         + "Nun.Reg. " + Format(mMesActivo, "00") + xNumAsiento, vbInformation, xTitulo
    
    MsgBox "El canje del Documento se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito " & vbCr & "Registro Nº: " & xNumAsiento, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        
    
    Grabar = True
    Exit Function
    
LaCague:
    'Resume
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el canje por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstCab = Nothing
    Set RstDet = Nothing
''''    Set RstDia = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation, xTitulo
        Cancel = 1
        Exit Sub
    End If
    Set RstFrm = Nothing
    Set Dg1.DataSource = Nothing
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then
            MuestraSegundoTab
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstFrm.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
       
    If Button.Index = 8 Then Filtrar
   
   If Button.Index = 9 Then
        If RstFrm.State = 0 Then Exit Sub
        RstFrm.Filter = adFilterNone
        RstFrm.Requery
        Dg1.Refresh
    End If
    
    If Button.Index = 10 Then Buscar
    If Button.Index = 11 Then CambiarMes
    If Button.Index = 15 Then
        Set RstFrm = Nothing
        Unload Me
    End If

End Sub

Sub MuestraSegundoTab()
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay registro para modificar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    Blanquea
    Dim rst As New ADODB.Recordset
    Dim nSQL As String
    TxtFchEmi.Valor = NulosC(RstFrm("fchemi"))
    TxtFchEmi_Validate False
    TxtRucPro.Text = NulosC(RstFrm("rucpro"))
    LblProveedor.Caption = NulosC(RstFrm("nompro"))
    
    TxtRucCli.Text = NulosC(RstFrm("ruccli"))
    LblCliente.Caption = NulosC(RstFrm("nomcli"))
    
    LblIdProveedor.Caption = NulosN(RstFrm("idpro"))
    LblIdCliente.Caption = NulosN(RstFrm("idcli"))
    
    TxtNumSer.Text = NulosC(RstFrm("numser"))
    TxtNumDoc.Text = NulosC(RstFrm("numdoc"))
    
    TxtIdMon.Text = NulosN(RstFrm("monid"))
    TxtIdMon_Validate False
    
    'mostramos los documentos del proveedor que se canjearon
    nSQL = "SELECT DISTINCT  mae_documento.codsun, mae_documento.abrev, com_compras!numser+'-'+com_compras!numdoc AS numdoc, com_compras.fchdoc, com_compras.fchven, mae_moneda.simbolo, con_canjesdet.impdoc, con_canjesdet.saldo, con_canjesdet.iddoc, con_canjesdet.idcan, con_diario.idcue " _
        + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (con_canjesdet LEFT JOIN com_compras ON con_canjesdet.iddoc = com_compras.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) LEFT JOIN con_diario ON (con_canjesdet.iddoc = con_diario.iddocpro) AND (con_canjesdet.idcan = con_diario.idmov) " _
        + vbCr + " WHERE con_diario.idlib = 8 and con_canjesdet.tipo = 2 AND  con_canjesdet.idcan = " & RstFrm("id") & " ; "

    RST_Busq rst, nSQL, xCon
    Fg1.Rows = 1
    
    If rst.RecordCount <> 0 Then rst.MoveFirst
    Do While Not rst.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(rst("abrev"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(rst("numdoc"))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(rst("fchdoc"), "dd/mm/yy")
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(rst("fchven"), "dd/mm/yy")
        Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(rst("simbolo"))
        Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(rst("impdoc")), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(rst("saldo")), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosN(rst("iddoc"))
        Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosN(rst("idcue"))
        Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(NulosN(rst("saldo")), FORMAT_MONTO)
        rst.MoveNext
    Loop
    
    'mostramos los documentos emitidos que se cajearon
    Set rst = Nothing
    nSQL = "SELECT DISTINCT mae_documento.codsun,  mae_documento.abrev,con_canjesdet.iddoc, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdoc, vta_ventas.fchdoc, vta_ventas.fchven, mae_moneda.simbolo, con_canjesdet.impdoc, con_canjesdet.saldo, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoccan, con_canjesdet.impcan, con_canjesdet.impsalcan, con_canjesdet.tipo, con_canjesdet.idcan,con_diario.idcue " _
        + vbCr + " FROM ((((con_canjesdet LEFT JOIN vta_ventas ON con_canjesdet.iddoc = vta_ventas.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN com_compras ON con_canjesdet.iddoccan = com_compras.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN con_diario ON (con_canjesdet.iddoc = con_diario.iddocpro) AND (con_canjesdet.idcan = con_diario.idmov) " _
        + vbCr + " WHERE con_diario.idlib = 8 and con_canjesdet.tipo = 1 AND con_canjesdet.idcan=" & RstFrm("id") & "; "
       
    RST_Busq rst, nSQL, xCon
    
    Fg2.Rows = 1
    If rst.RecordCount <> 0 Then rst.MoveFirst
    Do While Not rst.EOF
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(rst("abrev"))
        Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosC(rst("numdoc"))
        Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(rst("fchdoc"), "dd/mm/yy")
        Fg2.TextMatrix(Fg2.Rows - 1, 4) = NulosC(rst("simbolo"))
        Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(NulosN(rst("impdoc")), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 6) = Format(NulosN(rst("saldo")), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 7) = NulosC(rst("numdoccan"))
        Fg2.TextMatrix(Fg2.Rows - 1, 8) = Format(NulosN(rst("impcan")), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 9) = Format(NulosN(rst("impsalcan")), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 10) = Format(NulosN(rst("saldo")) - NulosN(rst("impcan")), FORMAT_MONTO)
        
        Fg2.TextMatrix(Fg2.Rows - 1, 11) = NulosN(rst("iddoc"))
        Fg2.TextMatrix(Fg2.Rows - 1, 12) = Fg1.TextMatrix(Fg1.Rows - 1, 8)
        Fg2.TextMatrix(Fg2.Rows - 1, 13) = NulosN(rst("idcue"))
        Fg2.TextMatrix(Fg2.Rows - 1, 14) = Fg1.TextMatrix(Fg1.Rows - 1, 9)
        
        rst.MoveNext
    Loop
    
    Set rst = Nothing
    
    HallarTotales
    
    If Fg2.Rows > 1 Then
        With Fg2
            .Select 1, 1, Fg2.Rows - 1, 6
            .FillStyle = flexFillRepeat
            .CellBackColor = &HDBF8F9
        
            .Select 1, 10, Fg2.Rows - 1, 10
            .FillStyle = flexFillRepeat
            .CellBackColor = &HDBF8F9
            
            .Select 1, 1, 1, 1
        End With
    End If
End Sub

Sub Eliminar()
    Dim Rpta, A As Integer
    Dim rst As New ADODB.Recordset
    
    If RstFrm.RecordCount = 0 Or RstFrm.EOF = True Or RstFrm.BOF = True Then
        MsgBox "No hay registro para eliminar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    Rpta = MsgBox("Esta seguro de eliminar el canje seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        On Error GoTo error
        RST_Busq rst, "SELECT con_canjesdet.tipo, con_canjesdet.iddoc, con_canjesdet.iddoccan, con_canjesdet.impcan From con_canjesdet " _
            & " WHERE (((con_canjesdet.tipo)=1) AND ((con_canjesdet.idcan)=" & RstFrm("id") & "))", xCon
        
        If rst.RecordCount <> 0 Then
            xCon.BeginTrans
            rst.MoveFirst
            Do While Not rst.EOF
                'actualizamos el saldo del documento de venta y del documento de compra
                xCon.Execute "UPDATE com_compras SET com_compras.impsal = [com_compras]![impsal]+" & rst("impcan") & " WHERE (com_compras.id = " & rst("iddoccan") & " )"
                'actualizamos el saldo del documento de venta y del documento de compra
                xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = [vta_ventas]![impsal]+" & rst("impcan") & " WHERE (vta_ventas.id =" & rst("iddoc") & ")"
                
                rst.MoveNext
            Loop
            xCon.Execute "DELETE * FROM con_canjes WHERE id = " & RstFrm("id") & ";"
            'eliminamos los asientos contables
            xCon.Execute "DELETE * FROM con_diario WHERE idlib =  8 AND idmov = " & RstFrm("id") & ";"
            
            'Eliminar historial del registro
            xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstFrm("id") & " AND idform = " & IdMenuActivo

            xCon.CommitTrans
            MsgBox "El canje se eliminó con éxito", vbInformation + vbOKCancel + vbDefaultButton1, xTitulo
            RstFrm.Requery
            Dg1.Refresh
            TabOne1.CurrTab = 0
        Else
            MsgBox "El canje no tiene documentos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
        Set rst = Nothing
    End If
    Exit Sub
error:
    xCon.RollbackTrans
    SHOW_ERROR Me.Name, "Eliminar"
End Sub


Private Sub TxtFchEmi_Validate(Cancel As Boolean)
    If IsDate(TxtFchEmi.Valor) = True Then
        LblTipoCambio.Caption = HallaTipoCambio(TxtFchEmi.Valor, 2, Venta, xCon)
        If NulosN(LblTipoCambio.Caption) = 0 Then
            LblTipoCambio.Caption = "Falta Registrar..."
        End If
    Else
        LblTipoCambio.Caption = ""
    End If
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtIdMon_Validate(Cancel As Boolean)
    If NulosC(TxtIdMon.Text) = "" Then
        LblMoneda.Caption = ""
        Exit Sub
    End If

    LblMoneda.Caption = Busca_Codigo(TxtIdMon.Text, "id", "descripcion", "mae_moneda", "N", xCon)

    If LblMoneda.Caption = "" Then
        TxtIdMon.Text = ""
        LblMoneda.Caption = ""
    End If
    
End Sub


'************************************
Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub pCargarGrid()
    On Error GoTo error
    Dim nSQL As String
    LblPeriodo(0).Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    LblPeriodo(1).Caption = LblPeriodo(0).Caption
    
    '------------------------------------------------------------------------------------------
    '--bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    '------------------------------------------------------------------------------------------
    
    
    nSQL = "SELECT con_canjes.*, mae_prov.numruc AS rucpro, mae_prov.nombre AS nompro, mae_cliente.numruc AS ruccli, mae_cliente.nombre AS nomcli, con_canjes.numser & '-' & con_canjes.numdoc AS numerodoc, con_canjes.idmon AS monid, mae_moneda.simbolo AS monabrev " _
            + vbCr + " FROM ((con_canjes LEFT JOIN mae_prov ON con_canjes.idpro = mae_prov.id) LEFT JOIN mae_cliente ON con_canjes.idcli = mae_cliente.id) LEFT JOIN mae_moneda ON con_canjes.idmon = mae_moneda.id " _
            + vbCr + " WHERE (((con_canjes.ano) = " & AnoTra & ") And ((con_canjes.idmes) = " & mMesActivo & ")) " _
            + vbCr + " ORDER BY con_canjes.numreg"


    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQL, xCon
    
    Set Dg1.DataSource = RstFrm
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pCargarGrid"
End Sub

Private Sub CambiarMes()
    mMesActivo = SeleccionaMes(xCon)
    pCargarGrid
    TabOne1.CurrTab = 0
End Sub


Private Sub Filtrar()
    
    ReDim xCampos(5, 4) As String
    xCampos(0, 0) = "Num.Reg.":    xCampos(0, 1) = "numreg":    xCampos(0, 2) = "C":   xCampos(0, 3) = "800"
    xCampos(1, 0) = "Número Doc.": xCampos(1, 1) = "numerodoc": xCampos(1, 2) = "C":   xCampos(1, 3) = "1000"
    xCampos(2, 0) = "Cliente":     xCampos(2, 1) = "nomcli":    xCampos(2, 2) = "C":   xCampos(2, 3) = "3200"
    xCampos(3, 0) = "Proveedor":   xCampos(3, 1) = "nompro":    xCampos(3, 2) = "C":   xCampos(3, 3) = "1000"
    xCampos(4, 0) = "Fch.Emi":     xCampos(4, 1) = "fchemi":    xCampos(4, 2) = "F":   xCampos(4, 3) = "900"
    xCampos(5, 0) = "Importe":     xCampos(5, 1) = "impcan":    xCampos(5, 2) = "N":   xCampos(5, 3) = "1000"
    
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg1
    Me.TabOne1.CurrTab = 0
    
End Sub

Private Function fValidarDatos() As Boolean
    If NulosC(TxtFchEmi.Valor) = "" Or IsDate(TxtFchEmi.Valor) = False Then
        MsgBox "No ha especificado la fecha de emisión del documento", vbExclamation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchEmi.SetFocus
        Exit Function
    End If
    
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "No ha especificado la moneda de la operación", vbExclamation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMon.SetFocus
        Exit Function
    End If
    If Trim(TxtNumSer.Text) = "" Then
        MsgBox "No ha especificado el N° de Serie. ", vbExclamation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSer.SetFocus
        Exit Function
    End If
    If Trim(TxtNumDoc.Text) = "" Then
        MsgBox "No ha especificado el N° de Documento ", vbExclamation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDoc.SetFocus
        Exit Function
    End If
    
    If NulosN(LblIdProveedor.Caption) = 0 Then
        MsgBox "No ha especificado el nombre del Proveedor", vbExclamation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtRucPro.SetFocus
        Exit Function
    End If
    
    If NulosN(LblIdCliente.Caption) = 0 Then
        MsgBox "No ha especificado el nombre del Cliente", vbExclamation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtRucCli.SetFocus
        Exit Function
    End If
    
'    If NulosN(TxtTotal4.Text) <> NulosN(TxtTotal3.Text) Then
'        MsgBox "La operación de canje esta mal efectuada, el total importe de documentos del proveedor " + Chr(13) _
'            & "no coincide con el total abonos del cliente" + vbCr + _
'            "Importe Proveedor: " + Format(NulosN(TxtTotal4.Text), FORMAT_MONTO) + vbCr + _
'            "Importe Cliente:  " + Format(NulosN(TxtTotal3.Text), FORMAT_MONTO), vbExclamation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtTotal4.SetFocus
'        Exit Function
'    End If
        
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado los documentos a canjear", vbExclamation + vbOKOnly + vbDefaultButton1
        Fg1.SetFocus
        Exit Function
    End If

    'eliminamos los documentos que no hayan sido canjeados
    Dim A As Integer
    For A = 1 To Fg2.Rows - 1
        If NulosC(Fg2.TextMatrix(A, 6)) = "" Then
            Fg2.RemoveItem A
            A = A - 1
        End If
        If Fg2.Rows = 1 Then
            Exit For
        End If
    Next A
    
    If Fg2.Rows = 1 Then
        MsgBox "No ha especificado los documentos canjeados" + vbCr + "haga clic sobre el botón [Canjear Documentos] ", vbExclamation + vbOKOnly + vbDefaultButton1
        Fg2.SetFocus
        Exit Function
    End If
    
    fValidarDatos = True
    
End Function

Private Sub Buscar()
    On Error GoTo error
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    ReDim xCampos(6, 4) As String
    
    xCampos(0, 0) = "Num.Reg.":     xCampos(0, 1) = "numreg":       xCampos(0, 2) = "900":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Núm.Doc.":     xCampos(1, 1) = "numerodoc":    xCampos(1, 2) = "1000":        xCampos(1, 3) = "C"
    xCampos(2, 0) = "Fch.Emi":      xCampos(2, 1) = "emision":      xCampos(2, 2) = "900":         xCampos(2, 3) = "F"
    xCampos(3, 0) = "Cliente":      xCampos(3, 1) = "nomcli":       xCampos(3, 2) = "2200":        xCampos(3, 3) = "C"
    xCampos(4, 0) = "Proveedor":    xCampos(4, 1) = "nompro":       xCampos(4, 2) = "2200":        xCampos(4, 3) = "C"
    xCampos(5, 0) = "Importe":      xCampos(5, 1) = "impcan":       xCampos(5, 2) = "1000":        xCampos(5, 3) = "N"
            
    nSQL = "SELECT con_canjes.*, mae_prov.numruc AS rucpro, mae_prov.nombre AS nompro, mae_cliente.numruc AS ruccli, mae_cliente.nombre AS nomcli, con_canjes.numser & '-' & con_canjes.numdoc AS numerodoc, con_canjes.idmon AS monid, mae_moneda.simbolo AS monabrev,format(con_canjes.fchemi,'dd/mm/yy') as emision  " _
            + vbCr + " FROM ((con_canjes LEFT JOIN mae_prov ON con_canjes.idpro = mae_prov.id) LEFT JOIN mae_cliente ON con_canjes.idcli = mae_cliente.id) LEFT JOIN mae_moneda ON con_canjes.idmon = mae_moneda.id " _
            + vbCr + " WHERE (((con_canjes.ano) = " & AnoTra & ") And ((con_canjes.idmes) = " & mMesActivo & ")) " _
            + vbCr + " ORDER BY con_canjes.numreg"
            
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Canjes de Documento", "nomcli", "nomcli", Principio
    
    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir
    
    RstFrm.MoveFirst
    RstFrm.Find "id = " & xRs("id") & ""
Salir:
    Set xRs = Nothing
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumSer_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        
    End If
End Sub

Private Sub TxtNumSer_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtNumSer.Text) = "" Then Exit Sub
    TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtNumDoc.Text) = "" Then Exit Sub
    TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
End Sub
'---PROVEEDOR
Private Sub TxtRucPro_Change()
    If Trim(TxtRucPro) = "" Then
        LblProveedor.Caption = ""
        LblIdProveedor.Caption = ""
        Fg1.Rows = 1
        Fg2.Rows = 1
        HallarTotales
    End If
End Sub

Private Sub TxtRucPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtRucPro_Validate True
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtRucPro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then CmdBusProv_Click
End Sub

Private Sub TxtRucPro_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If TxtRucPro.Text <> "" Then
        Dim rst As New ADODB.Recordset
        LblIdProveedor.Tag = LblIdProveedor.Caption
        RST_Busq rst, "SELECT * FROM mae_prov WHERE numruc = '" & Trim(TxtRucPro.Text) & "'", xCon
        If rst.RecordCount <> 0 Then
            rst.MoveFirst
            TxtRucPro.Text = NulosC(rst("numruc"))
            LblProveedor.Caption = NulosC(rst("nombre"))
            LblIdProveedor.Caption = NulosN(rst("id"))
        Else
            TxtRucPro.Text = ""
        End If
        Set rst = Nothing
        '----------------
        If NulosN(LblIdProveedor.Tag) <> NulosN(LblIdProveedor.Caption) Then
            Fg1.Rows = 1
            Fg2.Rows = 1
            HallarTotales
        End If
        '----------------
    End If
End Sub

'---CLIENTE
Private Sub TxtRucCli_Change()
    If Trim(TxtRucCli.Text) = "" Then
        LblCliente.Caption = ""
        LblIdCliente.Caption = ""
        Fg2.Rows = 1
        HallarTotales
    End If
End Sub

Private Sub TxtRucCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtRucCli_Validate True
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtRucCli_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then CmdBusCli_Click
End Sub

Private Sub TxtRucCli_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If TxtRucCli.Text <> "" Then
        Dim rst As New ADODB.Recordset
        LblIdCliente.Tag = LblIdCliente.Caption
        RST_Busq rst, "SELECT * FROM mae_cliente WHERE numruc = '" & Trim(TxtRucCli.Text) & "'", xCon
        If rst.RecordCount <> 0 Then
            rst.MoveFirst
            TxtRucCli.Text = NulosC(rst("numruc"))
            LblCliente.Caption = NulosC(rst("nombre"))
            LblIdCliente.Caption = NulosN(rst("id"))
        Else
            TxtRucCli.Text = ""
        End If
        Set rst = Nothing
        '----------------
        If NulosN(LblIdCliente.Tag) <> NulosN(LblIdCliente.Caption) Then
            Fg2.Rows = 1
            HallarTotales
        End If
        '----------------
    End If
End Sub
