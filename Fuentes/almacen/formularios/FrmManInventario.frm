VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManInventario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Almacen - Inventario"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame frm 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Index           =   5
      Left            =   30
      TabIndex        =   75
      Top             =   7650
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   225
         TabIndex        =   76
         Top             =   465
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
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
         Height          =   165
         Index           =   32
         Left            =   225
         TabIndex        =   78
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label lblProcesado 
         Alignment       =   2  'Center
         Caption         =   "lblProcesado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1440
         TabIndex        =   77
         Top             =   180
         Width           =   4260
      End
      Begin VB.Shape Shape1 
         Height          =   765
         Index           =   0
         Left            =   60
         Top             =   90
         Width           =   5805
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   0
         X1              =   5930
         X2              =   5930
         Y1              =   0
         Y2              =   945
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   5940
         Y1              =   935
         Y2              =   935
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
      Height          =   4320
      Index           =   0
      Left            =   90
      TabIndex        =   69
      Top             =   13050
      Visible         =   0   'False
      Width           =   11030
      Begin VB.CommandButton Cmd 
         Caption         =   "Aceptar"
         Height          =   330
         Index           =   12
         Left            =   100
         TabIndex        =   71
         ToolTipText     =   "Eliminar Todos"
         Top             =   3900
         Width           =   1155
      End
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   10710
         Picture         =   "FrmManInventario.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   70
         ToolTipText     =   "Cerrar"
         Top             =   45
         Width           =   195
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   3345
         Index           =   3
         Left            =   90
         TabIndex        =   73
         Top             =   420
         Width           =   10785
         _cx             =   19024
         _cy             =   5900
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
         Rows            =   3
         Cols            =   13
         FixedRows       =   2
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmManInventario.frx":02EC
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuadre de Stocks"
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
         Index           =   3
         Left            =   105
         TabIndex        =   72
         Top             =   60
         Width           =   1530
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   2
         X1              =   30
         X2              =   11000
         Y1              =   4290
         Y2              =   4290
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   2
         X1              =   11000
         X2              =   11000
         Y1              =   0
         Y2              =   4290
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   30
         Top             =   30
         Width           =   10920
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
      Height          =   4320
      Index           =   4
      Left            =   90
      TabIndex        =   49
      Top             =   8640
      Visible         =   0   'False
      Width           =   7530
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   3
         Left            =   7250
         Picture         =   "FrmManInventario.frx":04CE
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   62
         ToolTipText     =   "Cerrar"
         Top             =   45
         Width           =   195
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "Aceptar"
         Height          =   330
         Index           =   10
         Left            =   100
         TabIndex        =   61
         ToolTipText     =   "Eliminar Todos"
         Top             =   3900
         Width           =   1155
      End
      Begin SizerOneLibCtl.TabOne TabOne2 
         Height          =   3495
         Left            =   120
         TabIndex        =   50
         Top             =   360
         Width           =   7215
         _cx             =   12726
         _cy             =   6165
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         Caption         =   "  &Ingresos  |   &Salidas   "
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
         Begin VB.Frame Frame16 
            BorderStyle     =   0  'None
            Height          =   3180
            Left            =   7830
            TabIndex        =   56
            Top             =   300
            Width           =   7185
            Begin VB.TextBox NumMovSalidaText 
               Height          =   285
               Left            =   1080
               TabIndex        =   57
               Text            =   "NumMovSalidaText"
               Top             =   120
               Width           =   2535
            End
            Begin VSFlex7Ctl.VSFlexGrid fg 
               Height          =   2400
               Index           =   2
               Left            =   120
               TabIndex        =   58
               Top             =   600
               Width           =   6885
               _cx             =   12144
               _cy             =   4233
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
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManInventario.frx":07BA
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
            Begin VB.Label IdMovSalidaLabel 
               AutoSize        =   -1  'True
               Caption         =   "IdMovSalidaLabel"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   4305
               TabIndex        =   60
               Top             =   120
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "Num.Mov."
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   59
               Top             =   120
               Width           =   735
            End
         End
         Begin VB.Frame Frame15 
            BorderStyle     =   0  'None
            Height          =   3180
            Left            =   15
            TabIndex        =   51
            Top             =   300
            Width           =   7185
            Begin VB.TextBox NumMovIngresoText 
               Height          =   285
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   52
               Text            =   "NumMovIngresoText"
               Top             =   120
               Width           =   2535
            End
            Begin VSFlex7Ctl.VSFlexGrid fg 
               Height          =   2400
               Index           =   1
               Left            =   120
               TabIndex        =   53
               Top             =   600
               Width           =   6885
               _cx             =   12144
               _cy             =   4233
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
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManInventario.frx":086C
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
            Begin VB.Label IdMovIngresoLabel 
               AutoSize        =   -1  'True
               Caption         =   "IdMovIngresoLabel"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   4290
               TabIndex        =   55
               Top             =   180
               Visible         =   0   'False
               Width           =   1365
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "Num.Mov."
               Height          =   195
               Index           =   18
               Left            =   120
               TabIndex        =   54
               Top             =   120
               Width           =   735
            End
         End
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   5
         X1              =   7500
         X2              =   7500
         Y1              =   0
         Y2              =   4290
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   4
         X1              =   0
         X2              =   8295
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   5
         X1              =   30
         X2              =   7500
         Y1              =   4290
         Y2              =   4290
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resumen de Movimientos"
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
         TabIndex        =   63
         Top             =   60
         Width           =   2175
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   30
         Top             =   30
         Width           =   7440
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
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
            Picture         =   "FrmManInventario.frx":0918
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManInventario.frx":0E5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManInventario.frx":11EE
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManInventario.frx":1372
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManInventario.frx":17C6
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManInventario.frx":18DE
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManInventario.frx":1E22
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManInventario.frx":2366
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManInventario.frx":247A
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManInventario.frx":258E
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManInventario.frx":29E2
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManInventario.frx":2B4E
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManInventario.frx":3096
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   0
      TabIndex        =   24
      Top             =   375
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
         Caption         =   "Frame1"
         Height          =   6795
         Left            =   12525
         TabIndex        =   27
         Top             =   375
         Width           =   11790
         Begin MSComDlg.CommonDialog Cmm 
            Left            =   11280
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Frame FrmReceta 
            Caption         =   "[ Inventario ]"
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
            Height          =   1875
            Left            =   60
            TabIndex        =   33
            Top             =   390
            Width           =   11640
            Begin VB.CommandButton Cmd 
               Enabled         =   0   'False
               Height          =   240
               Index           =   2
               Left            =   1995
               Picture         =   "FrmManInventario.frx":3428
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   1110
               Width           =   240
            End
            Begin VB.TextBox DescripcionTextBox 
               Height          =   300
               Left            =   1350
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   5
               Text            =   "DescripcionTextBox"
               Top             =   1440
               Width           =   4665
            End
            Begin VB.TextBox NumDocText 
               Height          =   300
               Left            =   2450
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   1
               Top             =   360
               Width           =   3570
            End
            Begin VB.TextBox NumSerText 
               Height          =   300
               Left            =   1350
               Locked          =   -1  'True
               MaxLength       =   4
               TabIndex        =   0
               Top             =   360
               Width           =   915
            End
            Begin VB.Frame Frame3 
               Height          =   1080
               Left            =   6200
               TabIndex        =   42
               Top             =   700
               Width           =   5320
               Begin VB.CommandButton AnularButton 
                  Caption         =   "&Anular"
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   2520
                  TabIndex        =   22
                  Top             =   690
                  Width           =   1200
               End
               Begin VB.CommandButton AprobarButton 
                  Caption         =   "Aprobar"
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   120
                  TabIndex        =   20
                  Top             =   690
                  Width           =   1200
               End
               Begin VB.CommandButton RechazarButton 
                  Caption         =   "Rechazar"
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1320
                  TabIndex        =   21
                  Top             =   690
                  Width           =   1200
               End
               Begin VB.Label IdEstadoLabel 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "IdEstadoLabel"
                  ForeColor       =   &H000000C0&
                  Height          =   195
                  Left            =   3960
                  TabIndex        =   48
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   1020
               End
               Begin VB.Label EstadoLabel 
                  Alignment       =   2  'Center
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Pendiente"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   450
                  Left            =   120
                  TabIndex        =   44
                  Top             =   195
                  Width           =   5055
               End
               Begin VB.Label LblIdEstado 
                  AutoSize        =   -1  'True
                  Caption         =   "LblIdEstado"
                  ForeColor       =   &H000000C0&
                  Height          =   195
                  Left            =   255
                  TabIndex        =   43
                  Top             =   315
                  Visible         =   0   'False
                  Width           =   840
               End
            End
            Begin VB.TextBox InventarioTextBox 
               Height          =   300
               Left            =   1350
               Locked          =   -1  'True
               TabIndex        =   2
               Text            =   "InventarioTextBox"
               Top             =   705
               Width           =   4665
            End
            Begin AspaTextBoxFecha.TextBoxFecha FchVigTextBoxFecha 
               Height          =   300
               Left            =   10200
               TabIndex        =   7
               Top             =   345
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
               Valor           =   "18/09/2007"
            End
            Begin AspaTextBoxFecha.TextBoxFecha FchInvTextBoxFecha 
               Height          =   300
               Left            =   7440
               TabIndex        =   6
               Top             =   345
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
               Valor           =   "18/09/2007"
            End
            Begin VB.TextBox IdResponsableText 
               Height          =   300
               Left            =   1350
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   3
               Text            =   "IdResponsableText"
               Top             =   1080
               Width           =   915
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Responsable"
               Height          =   195
               Index           =   5
               Left            =   90
               TabIndex        =   67
               Top             =   1110
               Width           =   930
            End
            Begin VB.Label ResponsableLabel 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "ResponsableLabel"
               ForeColor       =   &H00800000&
               Height          =   300
               Left            =   2295
               TabIndex        =   66
               Top             =   1080
               Width           =   3720
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Glosa"
               Height          =   195
               Index           =   9
               Left            =   90
               TabIndex        =   65
               Top             =   1440
               Width           =   405
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Num. Doc."
               Height          =   195
               Index           =   2
               Left            =   90
               TabIndex        =   47
               Top             =   405
               Width           =   765
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H80000001&
               BackStyle       =   1  'Opaque
               Height          =   90
               Index           =   1
               Left            =   2305
               Top             =   480
               Width           =   105
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Fch. Vigencia"
               Height          =   195
               Index           =   3
               Left            =   9000
               TabIndex        =   41
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Fch. Inventario"
               Height          =   195
               Index           =   4
               Left            =   6200
               TabIndex        =   40
               Top             =   360
               Width           =   1065
            End
            Begin VB.Label IdInventarioLabel 
               AutoSize        =   -1  'True
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "IdInventarioLabel"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   10200
               TabIndex        =   35
               Top             =   105
               Visible         =   0   'False
               Width           =   1230
            End
            Begin VB.Label lbl_cb_capt 
               AutoSize        =   -1  'True
               Caption         =   "Descripcion"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   34
               Top             =   720
               Width           =   840
            End
         End
         Begin VB.Frame FrmLinea 
            Caption         =   "[ Filtro ]"
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
            Height          =   4485
            Left            =   60
            TabIndex        =   30
            Top             =   2280
            Width           =   11640
            Begin VB.CommandButton Cmd 
               Enabled         =   0   'False
               Height          =   240
               Index           =   1
               Left            =   3645
               Picture         =   "FrmManInventario.frx":355A
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   630
               Width           =   240
            End
            Begin VB.CommandButton Cmd 
               Caption         =   "&Cargar"
               Enabled         =   0   'False
               Height          =   330
               Index           =   4
               Left            =   10040
               TabIndex        =   19
               ToolTipText     =   "Establece como Principal la Linea Actual "
               Top             =   600
               Width           =   1400
            End
            Begin VB.CommandButton Cmd 
               Caption         =   "&Exportar Formato"
               Enabled         =   0   'False
               Height          =   330
               Index           =   3
               Left            =   10040
               TabIndex        =   14
               ToolTipText     =   "Establece como Principal la Linea Actual "
               Top             =   240
               Width           =   1400
            End
            Begin VB.CommandButton Cmd 
               Enabled         =   0   'False
               Height          =   240
               Index           =   0
               Left            =   3645
               Picture         =   "FrmManInventario.frx":368C
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   270
               Width           =   240
            End
            Begin VB.OptionButton OptTipo 
               Caption         =   "Con Movimiento"
               Enabled         =   0   'False
               Height          =   225
               Index           =   1
               Left            =   120
               TabIndex        =   9
               Top             =   500
               Width           =   1515
            End
            Begin VB.OptionButton OptTipo 
               Caption         =   "Todos"
               Enabled         =   0   'False
               Height          =   225
               Index           =   0
               Left            =   120
               TabIndex        =   8
               Top             =   270
               Width           =   885
            End
            Begin VB.Frame FrmLineaBot 
               Height          =   3435
               Left            =   10000
               TabIndex        =   31
               Top             =   960
               Width           =   1515
               Begin VB.CommandButton Cmd 
                  Caption         =   "&Ver Cuadre"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   13
                  Left            =   60
                  TabIndex        =   74
                  ToolTipText     =   "Establece como Principal la Linea Actual "
                  Top             =   2970
                  Width           =   1400
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "&Buscar"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   11
                  Left            =   40
                  TabIndex        =   68
                  ToolTipText     =   "Establece como Principal la Linea Actual "
                  Top             =   200
                  Width           =   1400
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "&Ver Movimientos"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   9
                  Left            =   60
                  TabIndex        =   64
                  ToolTipText     =   "Establece como Principal la Linea Actual "
                  Top             =   2580
                  Width           =   1400
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "Eliminar &Todos"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   8
                  Left            =   40
                  TabIndex        =   18
                  ToolTipText     =   "Procesa los valores de Linea"
                  Top             =   1860
                  Width           =   1400
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "&Eliminar"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   7
                  Left            =   40
                  TabIndex        =   17
                  ToolTipText     =   "Procesa los valores de Linea"
                  Top             =   1470
                  Width           =   1400
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "&Seleccionar"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   6
                  Left            =   40
                  TabIndex        =   16
                  ToolTipText     =   "Carga los valores correspondientes en la Receta"
                  Top             =   1100
                  Width           =   1400
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "&Agregar"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   5
                  Left            =   40
                  TabIndex        =   15
                  ToolTipText     =   "Establece como Principal la Linea Actual "
                  Top             =   700
                  Width           =   1400
               End
               Begin VB.Label lblnItm 
                  AutoSize        =   -1  'True
                  Caption         =   "lblnItm"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Left            =   840
                  TabIndex        =   39
                  Top             =   2280
                  Width           =   570
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "# Ítems :"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   38
                  Top             =   2280
                  Width           =   615
               End
            End
            Begin VSFlex7Ctl.VSFlexGrid fg 
               Height          =   3345
               Index           =   0
               Left            =   120
               TabIndex        =   32
               Top             =   1065
               Width           =   9795
               _cx             =   17277
               _cy             =   5900
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
               Rows            =   3
               Cols            =   13
               FixedRows       =   2
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManInventario.frx":37BE
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
            Begin VB.TextBox IdAlmacenTextBox 
               Height          =   300
               Left            =   3000
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   10
               Text            =   "IdAlmacenTextBox"
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox IdTipoInventarioTextBox 
               Height          =   300
               Left            =   3000
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   12
               Text            =   "IdTipoInventarioTextBox"
               Top             =   600
               Width           =   915
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Inventario"
               Height          =   195
               Index           =   1
               Left            =   1800
               TabIndex        =   46
               Top             =   630
               Width           =   1065
            End
            Begin VB.Label TipoInventarioLabel 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TipoInventarioLabel"
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
               Left            =   3960
               TabIndex        =   45
               Top             =   600
               Width           =   5925
            End
            Begin VB.Line Line2 
               BorderColor     =   &H80000003&
               Index           =   0
               X1              =   1680
               X2              =   1680
               Y1              =   240
               Y2              =   880
            End
            Begin VB.Label AlmacenLabel 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "AlmacenLabel"
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
               Left            =   3960
               TabIndex        =   37
               Top             =   240
               Width           =   5925
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Almacén"
               Height          =   195
               Index           =   0
               Left            =   1800
               TabIndex        =   36
               Top             =   270
               Width           =   615
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Inventario"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   105
            TabIndex        =   28
            Top             =   30
            Width           =   11610
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6795
         Left            =   45
         TabIndex        =   23
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   30
            TabIndex        =   25
            Top             =   345
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11324
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Id"
            Columns(0).DataField=   "idtomainventario"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nombre"
            Columns(1).DataField=   "nombre"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Tipo"
            Columns(2).DataField=   "tipoinventario"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Almacén"
            Columns(3).DataField=   "almacen"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fech.Vigencia"
            Columns(4).DataField=   "fchvig"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Estado"
            Columns(5).DataField=   "estadoinventario"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=7646"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=7567"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=2566"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2487"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=4445"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=4366"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=2302"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2223"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=2566"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2487"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
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
            HeadLines       =   1
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
            _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=70,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(64)  =   ":id=34,.wraptext=-1"
            _StyleDefs(65)  =   "Named:id=35:Footing"
            _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(67)  =   "Named:id=36:Selected"
            _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=37:Caption"
            _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(71)  =   "Named:id=38:HighlightRow"
            _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=39:EvenRow"
            _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(75)  =   "Named:id=40:OddRow"
            _StyleDefs(76)  =   ":id=40,.parent=33"
            _StyleDefs(77)  =   "Named:id=41:RecordSelector"
            _StyleDefs(78)  =   ":id=41,.parent=34"
            _StyleDefs(79)  =   "Named:id=42:FilterBar"
            _StyleDefs(80)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Inventario"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   26
            Top             =   30
            Width           =   11610
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1058
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
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Item"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Activar Item"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar un Item"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Retirar Item"
               EndProperty
            EndProperty
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir "
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "Eliminar             "
      End
   End
End
Attribute VB_Name = "FrmManInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMANALAMACEN.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : AQUI SE CREAN, MODIFICAN Y ELIMINAN LOS ITEMS Y SE LES ASIGNA LA CUENTA CONTABLE.
'*                  : Y EL CENTRO DE COSTO.
'* DISEÑADO POR     : Jose Chacon
'* ULTIMA REVISION  : 24/04/2012
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim RstPro As New ADODB.Recordset
Dim xHorIni As Date
Dim fOrdenLista As Boolean
Dim IdMenuActivo As Integer
Dim mIdRegistro&
Dim cSQL As String
Dim Agregando As Boolean
Dim OrigFX As Long
Dim OrigFY As Long
Dim F As New SistemaLogica.Funciones

Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Bloquea()
    NumSerText.Locked = Not NumSerText.Locked
    NumDocText.Locked = Not NumDocText.Locked
    InventarioTextBox.Locked = Not InventarioTextBox.Locked
    DescripcionTextBox.Locked = Not DescripcionTextBox.Locked
    FchInvTextBoxFecha.Locked = Not FchInvTextBoxFecha.Locked
    FchVigTextBoxFecha.Locked = Not FchVigTextBoxFecha.Locked
    
    OptTipo(0).Enabled = Not OptTipo(0).Enabled
    OptTipo(1).Enabled = Not OptTipo(1).Enabled
    
    habilitar cmd, Not cmd(0).Enabled
    
    bloquearControles
End Sub

Private Sub cargaStockCosto(Optional Indice As Long = 0)
On Error GoTo BloqueError
    Dim A As Integer
    
    If Indice > 0 Then
        With fg(0)
            If (NulosN(IdTipoInventarioTextBox.Text) = NulosN(F.KeyValue("InventarioInicial", xCon))) Then
                .TextMatrix(Indice, .ColIndex("STOCKACT")) = Format(F.NuloNumeric(F.SaldoInicial(.TextMatrix(Indice, .ColIndex("IDITEM")), F.NuloNumeric(IdAlmacenTextBox.Text), xCon)), FORMAT_CANTIDAD)
                .TextMatrix(Indice, .ColIndex("PREUNIACT")) = Format(F.CostoInicial(F.NuloNumeric(.TextMatrix(Indice, .ColIndex("IDITEM"))), F.NuloNumeric(IdAlmacenTextBox.Text), xCon), FORMAT_MONTO)
            ElseIf (NulosN(IdTipoInventarioTextBox.Text) = NulosN(F.KeyValue("InventarioAjuste", xCon))) Then
                .TextMatrix(Indice, .ColIndex("STOCKACT")) = Format(F.SaldoActual(F.NuloNumeric(.TextMatrix(Indice, .ColIndex("IDITEM"))), F.NuloNumeric(IdAlmacenTextBox.Text), "01/01/" & AnoTra, FchInvTextBoxFecha.Valor, xCon), FORMAT_CANTIDAD)
                .TextMatrix(Indice, .ColIndex("PREUNIACT")) = Format(F.CostoActual(F.NuloNumeric(.TextMatrix(Indice, .ColIndex("IDITEM"))), F.NuloNumeric(IdAlmacenTextBox.Text), "01/01/" & AnoTra, FchInvTextBoxFecha.Valor, xCon), FORMAT_MONTO)
            End If
            .TextMatrix(Indice, .ColIndex("DIFCANTIDAD")) = Format(NulosN(.TextMatrix(Indice, .ColIndex("CANTIDAD"))) - NulosN(.TextMatrix(Indice, .ColIndex("STOCKACT"))), FORMAT_CANTIDAD)
            .TextMatrix(Indice, .ColIndex("DIFPREUNI")) = Format(NulosN(.TextMatrix(Indice, .ColIndex("PREUNI"))) - NulosN(.TextMatrix(Indice, .ColIndex("PREUNIACT"))), FORMAT_CANTIDAD)
        End With
    Else
        CentrarFrm frm(5)
        frm(5).Visible = True
        lblProcesado.Caption = ""
        lbl(32).Caption = "Actualizando:"
        PgBar.Min = 0
        PgBar.Max = fg(0).Rows - 1
        PgBar.Value = 0
        Agregando = True
        For A = fg(0).FixedRows To fg(0).Rows - 1
            DoEvents
            Me.Refresh
            frm(5).Refresh
            lblProcesado.Caption = NulosC(fg(0).TextMatrix(A, fg(0).ColIndex("ITEM")))
            PgBar.Value = A
            If (NulosN(IdTipoInventarioTextBox.Text) = NulosN(F.KeyValue("InventarioInicial", xCon))) Then
                fg(0).TextMatrix(A, fg(0).ColIndex("STOCKACT")) = Format(F.NuloNumeric(F.SaldoInicial(fg(0).TextMatrix(A, fg(0).ColIndex("IDITEM")), F.NuloNumeric(IdAlmacenTextBox.Text), xCon)), FORMAT_CANTIDAD)
                fg(0).TextMatrix(A, fg(0).ColIndex("PREUNIACT")) = Format(F.CostoInicial(F.NuloNumeric(fg(0).TextMatrix(A, fg(0).ColIndex("IDITEM"))), F.NuloNumeric(IdAlmacenTextBox.Text), xCon), FORMAT_MONTO)
            ElseIf (NulosN(IdTipoInventarioTextBox.Text) = NulosN(F.KeyValue("InventarioAjuste", xCon))) Then
                fg(0).TextMatrix(A, fg(0).ColIndex("STOCKACT")) = Format(F.SaldoActual(F.NuloNumeric(fg(0).TextMatrix(A, fg(0).ColIndex("IDITEM"))), F.NuloNumeric(IdAlmacenTextBox.Text), "01/01/" & AnoTra, FchInvTextBoxFecha.Valor, xCon), FORMAT_CANTIDAD)
                fg(0).TextMatrix(A, fg(0).ColIndex("PREUNIACT")) = Format(F.CostoActual(F.NuloNumeric(fg(0).TextMatrix(A, fg(0).ColIndex("IDITEM"))), F.NuloNumeric(IdAlmacenTextBox.Text), "01/01/" & AnoTra, FchInvTextBoxFecha.Valor, xCon), FORMAT_MONTO)
            End If
            
            fg(0).TextMatrix(A, fg(0).ColIndex("DIFCANTIDAD")) = Format(NulosN(fg(0).TextMatrix(A, fg(0).ColIndex("CANTIDAD"))) - NulosN(fg(0).TextMatrix(A, fg(0).ColIndex("STOCKACT"))), FORMAT_CANTIDAD)
            fg(0).TextMatrix(A, fg(0).ColIndex("DIFPREUNI")) = Format(NulosN(fg(0).TextMatrix(A, fg(0).ColIndex("PREUNI"))) - NulosN(fg(0).TextMatrix(A, fg(0).ColIndex("PREUNIACT"))), FORMAT_CANTIDAD)
        Next
        frm(5).Visible = False
        lblProcesado.Caption = ""
        Agregando = False
    End If
    Exit Sub
    
BloqueError:
    Agregando = False
    MsgBox "Ocurrio un error al cargar los valores: " & Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, "Mantenimiento Inventario - cargaStockCosto"
    frm(5).Visible = False
    lblProcesado.Caption = ""
End Sub

Private Sub bloquearControles()
    cmd(10).Enabled = True ' Boton Aceptar
    cmd(9).Enabled = False ' Boton Ver Movimientos
    If QueHace = 3 Then
        If NulosN(IdTipoInventarioTextBox.Text) = 2 Then ' Ajuste
            cmd(9).Enabled = True ' Boton Ver Movimientos
            cmd(13).Enabled = True ' Boton Ver Movimientos
        Else
            cmd(9).Enabled = False ' Boton Ver Movimientos
            cmd(13).Enabled = False ' Boton Ver Movimientos
        End If
        
        AprobarButton.Enabled = False
        RechazarButton.Enabled = False
        AnularButton.Enabled = False
        OptTipo(0).Enabled = False
        OptTipo(1).Enabled = False
        cmd(0).Enabled = False ' Boton Almacen
        cmd(1).Enabled = False ' Boton Tipo Inventario
        cmd(4).Enabled = False ' Boton Cargar
        cmd(5).Enabled = False ' Boton Agregar
        cmd(6).Enabled = False ' Boton Seleccionar
        cmd(7).Enabled = False ' Boton Eliminar
        cmd(8).Enabled = False ' Boton Eliminar Todos
    Else
        If (NulosN(IdEstadoLabel.Caption) = 1) Then ' Pendiente
            AprobarButton.Enabled = True
            RechazarButton.Enabled = True
            AnularButton.Enabled = False
            OptTipo(0).Enabled = True
            OptTipo(1).Enabled = True
            cmd(0).Enabled = True ' Boton Almacen
            cmd(1).Enabled = True ' Boton Tipo Inventario
            cmd(4).Enabled = True ' Boton Cargar
            cmd(5).Enabled = True ' Boton Agregar
            cmd(6).Enabled = True ' Boton Seleccionar
            cmd(7).Enabled = True ' Boton Eliminar
            cmd(8).Enabled = True ' Boton Eliminar Todos
            If QueHace = 1 Then
                'AprobarButton.Enabled = False
                RechazarButton.Enabled = False
                cmd(3).Enabled = True ' Boton Exportar
                cmd(4).Enabled = True ' Boton Cargar
            End If
        ElseIf (NulosN(IdEstadoLabel.Caption) = 2) Then ' Aprobado
            AprobarButton.Enabled = False
            RechazarButton.Enabled = False
            AnularButton.Enabled = True
            OptTipo(0).Enabled = False
            OptTipo(1).Enabled = False
            cmd(0).Enabled = False ' Boton Almacen
            cmd(1).Enabled = False ' Boton Tipo Inventario
            cmd(4).Enabled = True ' Boton Cargar
            cmd(5).Enabled = True ' Boton Agregar
            cmd(6).Enabled = True ' Boton Seleccionar
            cmd(7).Enabled = True ' Boton Eliminar
            cmd(8).Enabled = True ' Boton Eliminar Todos
        Else ' Rechazado o Anulado
            AprobarButton.Enabled = False
            RechazarButton.Enabled = False
            AnularButton.Enabled = False
            OptTipo(0).Enabled = False
            OptTipo(1).Enabled = False
            cmd(0).Enabled = False ' Boton Almacen
            cmd(1).Enabled = False ' Boton Tipo Inventario
            cmd(2).Enabled = False ' Boton Buscar
            cmd(4).Enabled = False ' Boton Cargar
            cmd(5).Enabled = False ' Boton Agregar
            cmd(6).Enabled = False ' Boton Seleccionar
            cmd(7).Enabled = False ' Boton Eliminar
            cmd(8).Enabled = False ' Boton Eliminar Todos
        End If
    End If
End Sub

Private Sub pConfigurarGrid(ByRef Grid As VSFlexGrid)
    Agregando = True
    If Grid.FixedRows = 0 Then Grid.Rows = 2: Grid.FixedRows = 2
    
    If NulosN(IdTipoInventarioTextBox.Text) = 1 Then ' Inicial
    
        GRID_COMBINAR Grid, 0, 1, 1, 1, "Codigo", flexAlignCenterCenter, False, flexMergeFixedOnly, , &H8000000F, False
        GRID_COMBINAR Grid, 0, 2, 1, 2, "Ítem", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR Grid, 0, 3, 1, 3, "U.M.", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR Grid, 0, 4, 0, 6, "Stock Inicial", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR Grid, 1, 4, 1, 4, "Actual", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR Grid, 1, 5, 1, 5, "Carga", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR Grid, 1, 6, 1, 6, "Dif.", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR Grid, 0, 7, 0, 9, "Costo Inicial", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR Grid, 1, 7, 1, 7, "Actual", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR Grid, 1, 8, 1, 8, "Carga", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR Grid, 1, 9, 1, 9, "Dif.", flexAlignCenterCenter, False, , , &H8000000F, False
    
        Grid.MergeCells = flexMergeFixedOnly
        Grid.FillStyle = flexFillSingle
    ElseIf NulosN(IdTipoInventarioTextBox.Text) = 2 Then ' Ajuste
    
        GRID_COMBINAR Grid, 0, 1, 1, 1, "Codigo", flexAlignCenterCenter, False, flexMergeFixedOnly, , &H8000000F, False
        GRID_COMBINAR Grid, 0, 2, 1, 2, "Ítem", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR Grid, 0, 3, 1, 3, "U.M.", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR Grid, 0, 4, 0, 6, "Stock", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR Grid, 1, 4, 1, 4, "Actual", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR Grid, 1, 5, 1, 5, "Carga", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR Grid, 1, 6, 1, 6, "Dif.", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR Grid, 0, 7, 0, 9, "Costo", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR Grid, 1, 7, 1, 7, "Actual", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR Grid, 1, 8, 1, 8, "Carga", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR Grid, 1, 9, 1, 9, "Dif.", flexAlignCenterCenter, False, , , &H8000000F, False
            
        Grid.MergeCells = flexMergeFixedOnly
        Grid.FillStyle = flexFillSingle
    End If
    Agregando = False
End Sub

Private Sub iniciarCampos()
    TabOne1.CurrTab = 0
    
    
    fg(0).AllowUserResizing = flexResizeColumns
    fg(0).ExplorerBar = flexExSortShowAndMove
    fg(0).SelectionMode = flexSelectionByRow
    fg(0).BackColorSel = &H80&
    fg(0).ForeColorSel = &H80000005
    
    fg(0).ColWidth(10) = 0
    fg(0).ColWidth(11) = 0
       
    CaracteresNumericos = "0123456789." & Chr(8)
End Sub

Private Sub pCargarGrid()
    cSQL = "SELECT alm_tomainventario.idtomainventario, alm_tomainventario.idalm, alm_tomainventario.idtipoinventario, alm_tomainventario.idestadoinventario, alm_tomainventario.numser, alm_tomainventario.numdoc, alm_tomainventario.nombre, alm_tomainventario.descripcion, alm_tomainventario.fchinv, alm_tomainventario.fchvig, alm_tomainventario.tipofiltro, mae_tipoinventario.descripcion AS tipoinventario, alm_almacenes.descripcion AS almacen, mae_estadoinventario.descripcion AS estadoinventario, alm_tomainventario.idresponsable, pla_empleados.nombre AS responsable " _
        + vbCr + "FROM (((alm_tomainventario INNER JOIN alm_almacenes ON alm_tomainventario.idalm = alm_almacenes.id) INNER JOIN mae_tipoinventario ON alm_tomainventario.idtipoinventario = mae_tipoinventario.id) INNER JOIN mae_estadoinventario ON alm_tomainventario.idestadoinventario = mae_estadoinventario.id) LEFT JOIN pla_empleados ON alm_tomainventario.idresponsable = pla_empleados.id " _
        + vbCr + "ORDER BY alm_tomainventario.fchvig;"
        
    RST_Busq RstPro, cSQL, xCon

    Set Dg1.DataSource = RstPro
End Sub

Sub Blanquea()
    InventarioTextBox.Text = ""
    IdResponsableText.Text = ""
    ResponsableLabel.Caption = ""
    IdInventarioLabel.Caption = ""
    DescripcionTextBox.Text = ""
    NumSerText.Text = ""
    NumDocText.Text = ""
    IdAlmacenTextBox.Text = ""
    AlmacenLabel.Caption = ""
    IdTipoInventarioTextBox.Text = ""
    TipoInventarioLabel.Caption = ""
    EstadoLabel.Caption = "Pendiente"
    NumMovIngresoText.Text = ""
    IdMovIngresoLabel.Caption = ""
    NumMovSalidaText.Text = ""
    IdMovSalidaLabel.Caption = ""
    fg(0).Rows = fg(0).FixedRows
    fg(1).Rows = fg(1).FixedRows
    fg(2).Rows = fg(2).FixedRows
    OptTipo(0).Value = False
    OptTipo(1).Value = False
    lblnItm.Caption = ""
End Sub

Private Sub AnularButton_Click()
    Dim Rpta As Integer
    Dim mensaje As String
    
On Error GoTo error
    
    mensaje = "¿Esta seguro de anular el inventario?" _
        + vbCr + "Esto eliminarà los movimientos de ajustes generados para cuadrar el stock en el sistema"
        
    Rpta = MsgBox(mensaje, vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbNo Then Exit Sub
    
    ' Se cambia de estado el inventario
    IdEstadoLabel.Caption = F.NuloNumeric(F.KeyValue("EstadoAnuladoInventario", xCon))
    EstadoLabel.Caption = "Anulado"
    Exit Sub
    
error:
    F.MostrarMensajeError "Ocurriò un error al intentar aprobar el registro, " & Err.Description, ""
End Sub

Private Sub AprobarButton_Click()
    Dim Rpta As Integer
    Dim RptaActualizar As Integer
    Dim mensaje As String
    
On Error GoTo error
    
    If NulosN(IdTipoInventarioTextBox.Text) = 1 Then ' Inicial
        mensaje = "¿Esta seguro de aprobar el inventario?" _
        + vbCr + "Esto cargara el stock inicial en el sistema, este proceso no puede anularse"
    ElseIf NulosN(IdTipoInventarioTextBox.Text) = 2 Then ' Ajuste
        mensaje = "¿Esta seguro de aprobar el inventario?" _
        + vbCr + "Esto generará los movimientos de ajustes necesarios para cuadrar el stock en el sistema"
    Else
        Exit Sub
    End If
    
    Rpta = MsgBox(mensaje, vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbNo Then Exit Sub
    
    ' Se actualizan los datos
    If (MsgBox("¿Desea actualizar los valores iniciales?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo) = vbYes) Then cargaStockCosto
    ' Se cambia de estado el inventario
    IdEstadoLabel.Caption = F.NuloNumeric(F.KeyValue("EstadoAprobadoInventario", xCon))
    EstadoLabel.Caption = "Aprobado"
    Exit Sub
    
error:
    F.MostrarMensajeError "Ocurriò un error al intentar aprobar el registro, " & Err.Description, ""
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim xCampos() As String
    Dim nSQLId As String
    Dim xRs As New ADODB.Recordset
    Dim nTitulo As String
    Dim F As New SistemaLogica.Funciones
    Dim xform As New eps_librerias.FormSeleccion
    Dim A As Integer
    Dim mConsultaAux As String
    Dim mIdTipoAlmacen As String
    Dim mIdTipPro As String
    Dim mIdFam As String
    Dim mIdClas As String
    Dim mIdSubClas As String
    
    Select Case Index
        Case 0 ' Almacen
            ReDim xCampos(2, 4) As String
            
            xCampos(0, 0) = "Código":               xCampos(0, 1) = "id":               xCampos(0, 2) = "1000":     xCampos(0, 3) = "N":    xCampos(0, 4) = "N"
            xCampos(1, 0) = "Descripción":      xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "3500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
            
            cSQL = "SELECT alm_almacenes.id, alm_almacenes.descripcion " _
                + vbCr + "FROM alm_almacenes;"
                        
            nTitulo = "Buscando Almacenes"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            IdAlmacenTextBox.Text = NulosN(xRs("id"))
            AlmacenLabel.Caption = NulosC(xRs("descripcion"))
            fg(0).Rows = fg(0).FixedRows
            IdTipoInventarioTextBox.SetFocus
        
        Case 1 ' TipoInventario
            ReDim xCampos(2, 4) As String
            
            xCampos(0, 0) = "Código":               xCampos(0, 1) = "id":               xCampos(0, 2) = "1000":     xCampos(0, 3) = "N":    xCampos(0, 4) = "N"
            xCampos(1, 0) = "Descripción":      xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "3500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
            
            cSQL = "SELECT mae_tipoinventario.id, mae_tipoinventario.descripcion " _
                + vbCr + "FROM mae_tipoinventario;"
                        
            nTitulo = "Buscando Tipos de Inventario"
            
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            IdTipoInventarioTextBox.Text = NulosN(xRs("id"))
            TipoInventarioLabel.Caption = NulosC(xRs("descripcion"))
            cmd(5).SetFocus
            pConfigurarGrid fg(0)
        
        Case 2 ' Responsable
            ReDim xCampos(2, 4) As String
            
            'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Apellidos y Nombres":  xCampos(0, 1) = "apenom":     xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":               xCampos(1, 1) = "id":         xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
                      
            
            nTitulo = "Buscando Responsables"
            
            cSQL = "SELECT pla_empleados.nombre AS apenom, pla_empleados.id " _
                + vbCr + "FROM pla_empleados " _
                + vbCr + "ORDER BY pla_empleados.nombre;"
                        
            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                            "apenom", "apenom", Principio, ""
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            IdResponsableText.Text = NulosN(xRs("id"))
            ResponsableLabel.Caption = NulosC(xRs("apenom"))
            DescripcionTextBox.SetFocus
            
            Set xRs = Nothing
        
        Case 3 ' Exportar Formato
            If (OptTipo(0).Value = False And OptTipo(1).Value = False) Then
                MsgBox "Debe seleccionar un tipo de filtro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            If (NulosN(IdAlmacenTextBox.Text) = 0) Then
                MsgBox "Debe seleccionar un almacén", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                IdAlmacenTextBox.SetFocus
                Exit Sub
            End If
            If (NulosN(IdTipoInventarioTextBox.Text) = 0) Then
                MsgBox "Debe seleccionar un tipo de inventario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                IdTipoInventarioTextBox.SetFocus
                Exit Sub
            End If
            If fg(0).Rows <= fg(0).FixedRows Then
                MsgBox "Debe cargar datos en la grilla", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            fg(0).ColAlignment(fg(0).ColIndex("STOCKACT")) = flexAlignRightCenter
            fg(0).ColAlignment(fg(0).ColIndex("CANTIDAD")) = flexAlignRightCenter
            ExportarExcel fg(0)
        
        Case 4 ' Cargar Formato
            If (OptTipo(0).Value = False And OptTipo(1).Value = False) Then
                MsgBox "Debe seleccionar un tipo de filtro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            If (NulosN(IdAlmacenTextBox.Text) = 0) Then
                MsgBox "Debe seleccionar un almacén", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                IdAlmacenTextBox.SetFocus
                Exit Sub
            End If
            If (NulosN(IdTipoInventarioTextBox.Text) = 0) Then
                MsgBox "Debe seleccionar un tipo de inventario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                IdTipoInventarioTextBox.SetFocus
                Exit Sub
            End If
            ImportarExcel
        
        Case 5 ' Agregar
            AgregarItem
        
        Case 6 ' Seleccionar
            seleccionarItem
        
        Case 7 ' Eliminar
            If fg(0).Row < 1 Then
                MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                fg(0).SetFocus
                Exit Sub
            End If
            
            If fg(0).Rows = fg(0).FixedRows Then
                MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                fg(0).SetFocus
                Exit Sub
            End If
            
            If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
            fg(0).RemoveItem fg(0).Row
            
        Case 8 ' Eliminar Todos
            If MsgBox("¿Esta seguro de eliminar todos los registros?", _
                            vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
            
            fg(0).Rows = fg(0).FixedRows
        
        Case 9 ' Ver Movimientos
            CentrarFrm frm(4)
            frm(4).Visible = True
            cargarMovimientos
        
        Case 10 ' Aceptar Resumen Movimientos
            frm(4).Visible = False
            
        Case 11 ' Carga Datos a la grilla
            CargarValores
        
        Case 12 ' Ocultar Cuadre
            frm(0).Visible = False
            
        Case 13 ' Ver Cuadre
            CentrarFrm frm(0)
            frm(0).Visible = True
            CargarCuadre
        
    End Select
End Sub

Private Sub CargarValores()
    Dim nSQLId As String
    Dim xRs As New ADODB.Recordset
    
    If (NulosN(IdAlmacenTextBox.Text) = 0) Then
        MsgBox "Debe seleccionar un almacén", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        IdAlmacenTextBox.SetFocus
        Exit Sub
    End If
    If (NulosN(IdTipoInventarioTextBox.Text) = 0) Then
        MsgBox "Debe seleccionar un tipo de inventario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        IdTipoInventarioTextBox.SetFocus
        Exit Sub
    End If
    
    ' generar la lista de personal para no considerar en la lista
    nSQLId = GENERAR_SQL_ID(fg(0), fg(0).ColIndex("IDITEM"), " AND alm_inventario.id", "NOT IN", True)
        
    cSQL = "SELECT alm_inventario.id AS iditem, alm_inventario.descripcion AS desitem, alm_inventario.codpro AS codigo, alm_inventario.idunimed, mae_tipoproducto.descripcion AS tippro, mae_familia.descripcion AS familia, mae_clase.descripcion AS clase, mae_subclase.descripcion AS subclase " _
        + vbCr + "FROM alm_ingreso INNER JOIN (((((alm_inventario LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id) LEFT JOIN mae_clase ON alm_inventario.idclas = mae_clase.id) LEFT JOIN mae_subclase ON alm_inventario.idsubclas = mae_subclase.id) INNER JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) INNER JOIN alm_ingresodet ON alm_inventario.id = alm_ingresodet.iditem) ON alm_ingreso.id = alm_ingresodet.id " _
        + vbCr + "WHERE (((alm_ingreso.idalm) = " & IdAlmacenTextBox.Text & ") And ((alm_inventario.activo) = True) And ((alm_ingresodet.cantidad) > 0)) " & nSQLId _
        + vbCr + "GROUP BY alm_inventario.id, alm_inventario.descripcion, alm_inventario.codpro, alm_inventario.idunimed, mae_tipoproducto.descripcion, mae_familia.descripcion, mae_clase.descripcion, mae_subclase.descripcion " _
        + vbCr + "ORDER BY alm_inventario.codpro"
    
    RST_Busq xRs, cSQL, xCon
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
          
    Me.Refresh
    xRs.MoveFirst
    CentrarFrm frm(5)
    frm(5).Visible = True
    lblProcesado.Caption = ""
    lbl(32).Caption = "Procesando:"
    PgBar.Min = 0
    PgBar.Max = xRs.RecordCount
    PgBar.Value = 0
    Agregando = True
    While Not xRs.EOF
        DoEvents
        frm(5).Refresh
        lblProcesado.Caption = NulosC(xRs("desitem"))
        PgBar.Value = PgBar.Value + 1
        fg(0).Rows = fg(0).Rows + 1
        fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("ITEM")) = NulosC(xRs("desitem"))
        fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("CODIGO")) = NulosC(xRs("codigo"))
        fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("IDITEM")) = NulosN(xRs("iditem"))
        fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("IDUNIMED")) = NulosN(xRs("idunimed"))
        fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("UM")) = Busca_Codigo(NulosN(xRs("idunimed")), "id", "abrev", "mae_unidades", "N", xCon)
        fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("TIENEMOV")) = -1
        fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("CANTIDAD")) = Format(0, FORMAT_CANTIDAD)
        fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("PREUNI")) = Format(0, FORMAT_CANTIDAD)
        ' Se cargan los datos adicionales
        cargaStockCosto fg(0).Rows - 1
        
        ' Se valida si el item agregado tiene stock
        If F.NuloNumeric(fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("STOCKACT"))) = 0 Then
            fg(0).RemoveItem fg(0).Rows - 1
        End If
                
        xRs.MoveNext
    Wend
    Agregando = False
    frm(5).Visible = False
    
    lblProcesado.Caption = ""
    lblnItm.Caption = fg(0).Rows - fg(0).FixedRows
    
    Set xRs = Nothing
End Sub

Private Sub AgregarItem()
    Dim nSQLId As String
    Dim xRs As New ADODB.Recordset
    Dim nTitulo As String
    Dim xCampos(6, 4) As String
    
    xCampos(0, 0) = "Ítem":             xCampos(0, 1) = "desitem":      xCampos(0, 2) = "2800":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
    xCampos(1, 0) = "Codigo":           xCampos(1, 1) = "codigo":       xCampos(1, 2) = "900":      xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
    xCampos(2, 0) = "T. Prod.":         xCampos(2, 1) = "tippro":       xCampos(2, 2) = "1400":     xCampos(2, 3) = "C":    xCampos(2, 4) = "C"
    xCampos(3, 0) = "Familia":          xCampos(3, 1) = "familia":      xCampos(3, 2) = "1000":     xCampos(3, 3) = "C":    xCampos(3, 4) = "C"
    xCampos(4, 0) = "Clase":            xCampos(4, 1) = "clase":        xCampos(4, 2) = "1000":     xCampos(4, 3) = "C":    xCampos(4, 4) = "C"
    xCampos(5, 0) = "Sub Clase":        xCampos(5, 1) = "subclase":     xCampos(5, 2) = "1000":     xCampos(5, 3) = "C":    xCampos(5, 4) = "C"
    
    If (OptTipo(0).Value = False And OptTipo(1).Value = False) Then
        MsgBox "Debe seleccionar un tipo de filtro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    If (NulosN(IdAlmacenTextBox.Text) = 0) Then
        MsgBox "Debe seleccionar un almacén", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        IdAlmacenTextBox.SetFocus
        Exit Sub
    End If
    If (NulosN(IdTipoInventarioTextBox.Text) = 0) Then
        MsgBox "Debe seleccionar un tipo de inventario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        IdTipoInventarioTextBox.SetFocus
        Exit Sub
    End If
      
    ' generar la lista de personal para no considerar en la lista
    nSQLId = GENERAR_SQL_ID(fg(0), fg(0).ColIndex("IDITEM"), " AND alm_inventario.id", "NOT IN", True)
    
    If (OptTipo(1).Value = True) Then ' Con movimiento
        nSQLId = nSQLId & " AND alm_inventario.id IN (SELECT alm_ingresodet.iditem " _
                                    + vbCr + "FROM alm_ingresodet " _
                                    + vbCr + "WHERE (((alm_ingresodet.Cantidad) > 0)) " _
                                    + vbCr + "GROUP BY alm_ingresodet.iditem) "
    End If
    
    cSQL = "SELECT alm_inventario.id AS iditem, alm_inventario.descripcion AS desitem, alm_inventario.codpro AS codigo, alm_inventario.idunimed, mae_tipoproducto.descripcion AS tippro, mae_familia.descripcion AS familia, mae_clase.descripcion AS clase, mae_subclase.descripcion AS subclase " _
        + vbCr + "FROM (((alm_inventario LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id) LEFT JOIN mae_clase ON alm_inventario.idclas = mae_clase.id) LEFT JOIN mae_subclase ON alm_inventario.idsubclas = mae_subclase.id) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id " _
        + vbCr + "WHERE (((alm_inventario.activo)=-1)) " & nSQLId
                
    nTitulo = "Buscando Ítems"
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "codigo", "codigo", Principio
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    fg(0).Rows = fg(0).Rows + 1
    fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("ITEM")) = NulosC(xRs("desitem"))
    fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("CODIGO")) = NulosC(xRs("codigo"))
    fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("IDITEM")) = NulosN(xRs("iditem"))
    fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("IDUNIMED")) = NulosN(xRs("idunimed"))
    fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("UM")) = Busca_Codigo(NulosN(xRs("idunimed")), "id", "abrev", "mae_unidades", "N", xCon)
    fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("TIENEMOV")) = 0
    fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("CANTIDAD")) = Format(0, FORMAT_CANTIDAD)
    fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("PREUNI")) = Format(0, FORMAT_CANTIDAD)
    
    cargaStockCosto fg(0).Rows - 1

    fg(0).SetFocus
    fg(0).Row = fg(0).Rows - 1
    fg(0).Col = fg(0).ColIndex("CANTIDAD")
    lblnItm.Caption = fg(0).Rows - fg(0).FixedRows
    Set xRs = Nothing
End Sub

Private Sub seleccionarItem()
    Dim nSQLId As String
    Dim xRs As New ADODB.Recordset
    Dim nTitulo As String
    Dim xCampos(6, 4) As String
    
    xCampos(0, 0) = "Ítem":             xCampos(0, 1) = "desitem":      xCampos(0, 2) = "3500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
    xCampos(1, 0) = "Codigo":           xCampos(1, 1) = "codigo":       xCampos(1, 2) = "900":      xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
    xCampos(2, 0) = "T. Prod.":         xCampos(2, 1) = "tippro":       xCampos(2, 2) = "2000":     xCampos(2, 3) = "C":    xCampos(2, 4) = "C"
    xCampos(3, 0) = "Familia":          xCampos(3, 1) = "familia":      xCampos(3, 2) = "1000":     xCampos(3, 3) = "C":    xCampos(3, 4) = "C"
    xCampos(4, 0) = "Clase":            xCampos(4, 1) = "clase":        xCampos(4, 2) = "1200":     xCampos(4, 3) = "C":    xCampos(4, 4) = "C"
    xCampos(5, 0) = "Sub Clase":        xCampos(5, 1) = "subclase":     xCampos(5, 2) = "1000":     xCampos(5, 3) = "C":    xCampos(5, 4) = "C"
          
    If (OptTipo(0).Value = False And OptTipo(1).Value = False) Then
        MsgBox "Debe seleccionar un tipo de filtro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    If (NulosN(IdAlmacenTextBox.Text) = 0) Then
        MsgBox "Debe seleccionar un almacén", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        IdAlmacenTextBox.SetFocus
        Exit Sub
    End If
    If (NulosN(IdTipoInventarioTextBox.Text) = 0) Then
        MsgBox "Debe seleccionar un tipo de inventario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        IdTipoInventarioTextBox.SetFocus
        Exit Sub
    End If
    
    ' generar la lista de personal para no considerar en la lista
    nSQLId = GENERAR_SQL_ID(fg(0), fg(0).ColIndex("IDITEM"), " AND alm_inventario.id", "NOT IN", True)
    
    If (OptTipo(1).Value = True) Then ' Con movimiento
        nSQLId = nSQLId & " AND alm_inventario.id IN (SELECT alm_ingresodet.iditem " _
                                    + vbCr + "FROM alm_ingresodet " _
                                    + vbCr + "WHERE (((alm_ingresodet.Cantidad) > 0)) " _
                                    + vbCr + "GROUP BY alm_ingresodet.iditem) "
    End If
    
    cSQL = "SELECT 0 As xsel, alm_inventario.id AS iditem, alm_inventario.descripcion AS desitem, alm_inventario.codpro AS codigo, alm_inventario.idunimed, mae_tipoproducto.descripcion AS tippro, mae_familia.descripcion AS familia, mae_clase.descripcion AS clase, mae_subclase.descripcion AS subclase " _
        + vbCr + "FROM (((alm_inventario LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id) LEFT JOIN mae_clase ON alm_inventario.idclas = mae_clase.id) LEFT JOIN mae_subclase ON alm_inventario.idsubclas = mae_subclase.id) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id " _
        + vbCr + "WHERE (((alm_inventario.activo)=-1)) " & nSQLId _
        + vbCr + "ORDER BY alm_inventario.codpro"
    
    nTitulo = "Buscando Ítems"
    
    CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, cSQL, xCampos, nTitulo
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
          
    Me.Refresh
    xRs.MoveFirst
    CentrarFrm frm(5)
    frm(5).Visible = True
    lblProcesado.Caption = ""
    lbl(32).Caption = "Procesando:"
    PgBar.Min = 0
    PgBar.Max = xRs.RecordCount
    PgBar.Value = 0
    Agregando = True
    While Not xRs.EOF
        DoEvents
        frm(5).Refresh
        lblProcesado.Caption = NulosC(xRs("desitem"))
        PgBar.Value = PgBar.Value + 1
        fg(0).Rows = fg(0).Rows + 1
        fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("ITEM")) = NulosC(xRs("desitem"))
        fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("CODIGO")) = NulosC(xRs("codigo"))
        fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("IDITEM")) = NulosN(xRs("iditem"))
        fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("IDUNIMED")) = NulosN(xRs("idunimed"))
        fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("UM")) = Busca_Codigo(NulosN(xRs("idunimed")), "id", "abrev", "mae_unidades", "N", xCon)
        fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("TIENEMOV")) = 0
        fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("CANTIDAD")) = Format(0, FORMAT_CANTIDAD)
        fg(0).TextMatrix(fg(0).Rows - 1, fg(0).ColIndex("PREUNI")) = Format(0, FORMAT_CANTIDAD)
        ' Se cargan los datos adicionales
        cargaStockCosto fg(0).Rows - 1
                
        xRs.MoveNext
    Wend
    Agregando = False
    frm(5).Visible = False
    
    lblProcesado.Caption = ""
    lblnItm.Caption = fg(0).Rows - fg(0).FixedRows
    
    Set xRs = Nothing
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub cargarMovimientos()
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    
    NumMovIngresoText.Text = ""
    NumMovSalidaText.Text = ""
    
    cSQL = "SELECT alm_tomainventariomov.idtomainventario, alm_tomainventariomov.idingreso, alm_ingreso.tipmov, [alm_ingreso].[numser] & '-' & [alm_ingreso].[numdoc] AS numdoc, alm_ingresodet.iditem, alm_inventario.codpro, alm_inventario.descripcion As item, alm_ingresodet.cantidad " _
        + vbCr + "FROM (alm_tomainventariomov INNER JOIN alm_ingreso ON alm_tomainventariomov.idingreso = alm_ingreso.id) INNER JOIN (alm_ingresodet INNER JOIN alm_inventario ON alm_ingresodet.iditem = alm_inventario.id) ON alm_ingreso.id = alm_ingresodet.id " _
        + vbCr + "WHERE (alm_tomainventariomov.idtomainventario = " & NulosN(IdInventarioLabel.Caption) & ")"
        
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon

    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    fg(1).Rows = fg(1).FixedRows
    xRs.Filter = adFilterNone
    xRs.Filter = "tipmov=-1"
    If xRs.RecordCount = 0 Then GoTo SALIDAS
    xRs.MoveFirst
    NumMovIngresoText.Text = NulosC(xRs("numdoc"))
    IdMovIngresoLabel.Caption = NulosN(xRs("idingreso"))
    For A = 1 To xRs.RecordCount
        fg(1).Rows = fg(1).Rows + 1
        fg(1).TextMatrix(A, fg(1).ColIndex("IDITEM")) = NulosN(xRs("iditem"))
        fg(1).TextMatrix(A, fg(1).ColIndex("CANTIDAD")) = Format(NulosN(xRs("cantidad")), FORMAT_CANTIDAD)
        fg(1).TextMatrix(A, fg(1).ColIndex("CODIGO")) = NulosC(xRs("codpro"))
        fg(1).TextMatrix(A, fg(1).ColIndex("ITEM")) = NulosC(xRs("item"))
        xRs.MoveNext
        
        If xRs.EOF = True Then
            Exit For
        End If
    Next A
SALIDAS:
    fg(2).Rows = fg(2).FixedRows
    xRs.Filter = adFilterNone
    xRs.Filter = "tipmov=0"
    If xRs.RecordCount = 0 Then Exit Sub
    xRs.MoveFirst
    NumMovSalidaText.Text = NulosC(xRs("numdoc"))
    IdMovSalidaLabel.Caption = NulosN(xRs("idingreso"))
    For A = 1 To xRs.RecordCount
        fg(2).Rows = fg(2).Rows + 1
        fg(2).TextMatrix(A, fg(2).ColIndex("IDITEM")) = NulosN(xRs("iditem"))
        fg(2).TextMatrix(A, fg(2).ColIndex("CANTIDAD")) = Format(NulosN(xRs("cantidad")), FORMAT_CANTIDAD)
        fg(2).TextMatrix(A, fg(2).ColIndex("CODIGO")) = NulosC(xRs("codpro"))
        fg(2).TextMatrix(A, fg(2).ColIndex("ITEM")) = NulosC(xRs("item"))
        xRs.MoveNext
        
        If xRs.EOF = True Then
            Exit For
        End If
    Next A
End Sub

Private Sub CargarCuadre()
    Dim RstDet As New ADODB.Recordset
        
    cSQL = "SELECT alm_tomainventariodet.iditem, alm_tomainventariodet.idunimed, alm_tomainventariodet.stockactual, alm_tomainventariodet.preuniactual, alm_tomainventariodet.cantidad, alm_tomainventariodet.preuni, alm_inventario.codpro, alm_inventario.descripcion As item, mae_unidades.abrev As unimed " _
        + vbCr + "FROM (alm_tomainventariodet INNER JOIN alm_inventario ON alm_tomainventariodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON alm_tomainventariodet.idunimed = mae_unidades.id " _
        + vbCr + "WHERE (((alm_tomainventariodet.idtomainventario) = " & NulosN(RstPro("idtomainventario")) & "));"
    
    Set RstDet = Nothing
    RST_Busq RstDet, cSQL, xCon
    
    Agregando = True
    pConfigurarGrid fg(3)
    With fg(3)
        .Rows = .FixedRows
        If RstDet.RecordCount = 0 Then Exit Sub
        CentrarFrm frm(5)
        frm(5).Visible = True
        lblProcesado.Caption = ""
        lbl(32).Caption = "Cargando: "
        PgBar.Min = 0
        PgBar.Max = RstDet.RecordCount
        PgBar.Value = 0
        Agregando = True
        RstDet.MoveFirst
        While Not RstDet.EOF
            DoEvents
            .Rows = .Rows + 1
            frm(5).Refresh
            lblProcesado.Caption = NulosC(.TextMatrix(.Rows - 1, .ColIndex("ITEM")))
            PgBar.Value = .Rows - 2
            .TopRow = .Rows - 1
            .TextMatrix(.Rows - 1, .ColIndex("IDITEM")) = NulosN(RstDet("iditem"))
            .TextMatrix(.Rows - 1, .ColIndex("IDUNIMED")) = NulosN(RstDet("idunimed"))
            ' Datos actuales
            If (NulosN(IdTipoInventarioTextBox.Text) = NulosN(F.KeyValue("InventarioInicial", xCon))) Then
                .TextMatrix(.Rows - 1, .ColIndex("STOCKACT")) = Format(F.NuloNumeric(F.SaldoInicial(.TextMatrix(.Rows - 1, .ColIndex("IDITEM")), F.NuloNumeric(IdAlmacenTextBox.Text), xCon)), FORMAT_CANTIDAD)
                '.TextMatrix(.Rows - 1, .ColIndex("PREUNIACT")) = Format(F.CostoInicial(F.NuloNumeric(.TextMatrix(.Rows - 1, .ColIndex("IDITEM"))), F.NuloNumeric(IdAlmacenTextBox.Text), xCon), FORMAT_MONTO)
            ElseIf (NulosN(IdTipoInventarioTextBox.Text) = NulosN(F.KeyValue("InventarioAjuste", xCon))) Then
                .TextMatrix(.Rows - 1, .ColIndex("STOCKACT")) = Format(F.SaldoActual(F.NuloNumeric(.TextMatrix(.Rows - 1, .ColIndex("IDITEM"))), F.NuloNumeric(IdAlmacenTextBox.Text), "01/01/" & AnoTra, FchInvTextBoxFecha.Valor, xCon), FORMAT_CANTIDAD)
                '.TextMatrix(.Rows - 1, .ColIndex("PREUNIACT")) = Format(F.CostoActual(F.NuloNumeric(.TextMatrix(.Rows - 1, .ColIndex("IDITEM"))), F.NuloNumeric(IdAlmacenTextBox.Text), "01/01/" & AnoTra, FchInvTextBoxFecha.Valor, xCon), FORMAT_MONTO)
            End If
            ' Datos cargados
            .TextMatrix(.Rows - 1, .ColIndex("CANTIDAD")) = Format(NulosN(RstDet("cantidad")), FORMAT_CANTIDAD)
            .TextMatrix(.Rows - 1, .ColIndex("PREUNI")) = Format(NulosN(RstDet("preuni")), FORMAT_MONTO)
            ' Diferencias
            .TextMatrix(.Rows - 1, .ColIndex("DIFCANTIDAD")) = Format(NulosN(.TextMatrix(.Rows - 1, .ColIndex("CANTIDAD"))) - NulosN(.TextMatrix(.Rows - 1, .ColIndex("STOCKACT"))), FORMAT_CANTIDAD)
            '.TextMatrix(.Rows - 1, .ColIndex("DIFPREUNI")) = Format(NulosN(.TextMatrix(.Rows - 1, .ColIndex("PREUNI"))) - NulosN(.TextMatrix(.Rows - 1, .ColIndex("PREUNIACT"))), FORMAT_CANTIDAD)
            ' Datos adicionales
            .TextMatrix(.Rows - 1, .ColIndex("CODIGO")) = NulosC(RstDet("codpro"))
            .TextMatrix(.Rows - 1, .ColIndex("ITEM")) = NulosC(RstDet("item"))
            .TextMatrix(.Rows - 1, .ColIndex("UM")) = NulosC(RstDet("unimed"))
                        
            RstDet.MoveNext
        Wend
        ' Totales
        .Rows = .Rows + 1
        FORMATO_CELDA fg(3), .Rows - 1, .ColIndex("ITEM"), , True, , "TOTAL"
        .TextMatrix(.Rows - 1, .ColIndex("DIFCANTIDAD")) = Format(GRID_SUMAR_COL(fg(3), .ColIndex("DIFCANTIDAD")), FORMAT_MONTO)
        .TopRow = .Rows - 1
        
        ' Formato de Grilla
        .Select .FixedRows, .ColIndex("DIFCANTIDAD"), .Rows - 1, .ColIndex("DIFCANTIDAD")
        .FillStyle = flexFillRepeat
        .CellBackColor = &HC0FFFF
        F.PintarGrid fg(0), .ColIndex("DIFCANTIDAD"), &H0&, &HFF&
        .Select .Rows - 1, .ColIndex("DIFCANTIDAD")
                
        frm(5).Visible = False
        lblProcesado.Caption = ""
        Agregando = False
    End With
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstPro
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA LAS COLUMNAS DEL CONTROL Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstPro.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    ' SI SE HA PRESIONADO LA TECLA F12 MOSTRAMOS LA INFORMACION DE EDICION DEL REGISTRO
    If KeyCode = 123 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstPro("id")), xCon
    End If
End Sub

Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Row <= 0 Then Exit Sub
    
    If Index = 0 Then
        Select Case Col
            Case fg(0).ColIndex("CANTIDAD"), fg(0).ColIndex("PREUNI")
                fg(0).TextMatrix(Row, Col) = Format(F.NuloNumeric(fg(0).TextMatrix(Row, Col)), FORMAT_CANTIDAD)
                fg(0).TextMatrix(Row, fg(0).ColIndex("DIFCANTIDAD")) = Format(NulosN(fg(0).TextMatrix(Row, fg(0).ColIndex("CANTIDAD"))) - NulosN(fg(0).TextMatrix(Row, fg(0).ColIndex("STOCKACT"))), FORMAT_CANTIDAD)
                fg(0).TextMatrix(Row, fg(0).ColIndex("DIFPREUNI")) = Format(NulosN(fg(0).TextMatrix(Row, fg(0).ColIndex("PREUNI"))) - NulosN(fg(0).TextMatrix(Row, fg(0).ColIndex("PREUNIACT"))), FORMAT_CANTIDAD)
        End Select
    End If
End Sub

Private Sub fg_EnterCell(Index As Integer)
    If Index = 0 Then
        If QueHace = 3 Then fg(0).Editable = flexEDNone: Exit Sub
        
        Select Case fg(0).Col
            Case fg(0).ColIndex("CANTIDAD"), fg(0).ColIndex("PREUNI")
                fg(0).Editable = flexEDKbdMouse
                
            Case Else
                fg(0).Editable = flexEDNone
        End Select
    End If
End Sub

Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Index = 0 Then
        Select Case Col
            Case fg(0).ColIndex("CANTIDAD"), fg(0).ColIndex("PREUNI")
                If validar_numero(KeyAscii) = False Then KeyAscii = 0
        End Select
    End If
End Sub

Private Sub Form_Activate()
    ' CARGAMOS LOS ITEMS DEL INVENTARIO Y LOS MOSTRAMOS EN LA LA PRIMERA PESTAÑA DEL FORMULARIO, ESTE EVENTO SOLO SE EJECUTARA
    ' UNA SOLA VEZ
    If SeEjecuto = False Then
        Dim Rpta As Integer
        
        IdMenuActivo = xIdMenu
        
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        SeEjecuto = True
        pCargarGrid
        
    End If
End Sub

Sub Nuevo()
    QueHace = 1
    IdEstadoLabel.Caption = 1
    Blanquea
    Bloquea
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Agregando Inventario"
    
    fg(0).FixedRows = 0
    fg(0).Rows = fg(0).FixedRows
    fg(0).Editable = flexEDKbdMouse
    fg(0).SelectionMode = flexSelectionFree
    FchInvTextBoxFecha.Valor = Date
    FchVigTextBoxFecha.Valor = Date
    xHorIni = Time
    NumSerText.SetFocus
End Sub

Private Sub Form_Load()
    ' CARGAMOS EL FORMULARIO
    QueHace = 3
    SeEjecuto = False
    iniciarCampos
End Sub

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub

    If Me.Width <= 12000 Then Me.Width = 12000
    If Me.Height <= 4000 Then Me.Height = 4000

    ' Se dimensiona el Contenido
    TabOne1.Width = Me.Width - 90
    TabOne1.Height = Me.Height - 750
    
    Label4(0).Width = Me.Width - 100
    Dg1.Width = TabOne1.Width - 135
    Dg1.Height = TabOne1.Height - 795
    
    ' Se dimensiona el Detalle
    Label5.Width = Me.Width - 100
    
    FrmLinea.Width = TabOne1.Width - 240
    FrmLinea.Height = TabOne1.Height - 2730
    
    FrmLineaBot.Left = FrmLinea.Width - 1640
    FrmLineaBot.Height = FrmLinea.Height - 1050
    fg(0).Width = FrmLinea.Width - 1845
    fg(0).Height = FrmLinea.Height - 1140
End Sub

Private Sub NumSerText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub NumSerText_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    
    If NulosC(NumSerText.Text) <> "" Then
        NumSerText.Text = Format(NumSerText.Text, "0000")
        NumDocText.Text = hallarNumDoc("alm_tomainventario", "'" & NulosC(NumSerText.Text) & "'", "numser")
        If NulosC(NumDocText.Text) = "" Then NumSerText.Text = ""
    End If
End Sub

Private Sub PbCerrar_Click(Index As Integer)
    Select Case Index
        Case 0:
            frm(0).Visible = False
        Case 3:
            frm(4).Visible = False
    End Select
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 3 Then Exit Sub
        MuestraSegundoTab
    End If
    frm(0).Visible = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            RstPro.Requery
            Dg1.Refresh
            Cancelar
            
            If RstPro.RecordCount <> 0 Then
                RstPro.MoveFirst
                RstPro.Find "idtomainventario=" & mIdRegistro
                If RstPro.EOF = True Then RstPro.MoveFirst
            End If
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstPro.Filter = ""
    End If
    
    If Button.Index = 15 Then
        Set RstPro = Nothing
        Unload Me
    End If
End Sub

Function validarDatos() As Boolean
    ' Se valida la fecha de cierre de mes
    If F.MesCerradoOpcion(F.RetornarMesFecha(CDate(FchVigTextBoxFecha.Valor)), CLng(F.KeyValue("IdOpcionSistemaMovimientoAlmacen", xCon)), xCon) Then
        MsgBox "El mes al que pertenece la fecha de vigencia se encuentra cerrado, modifique la fecha o aperture el mes para la opcion: [Ingresos y Salidas de almacen]", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        FchVigTextBoxFecha.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosC(NumSerText.Text) = "" Then
        MsgBox "No ha especificado un nùmero de serie", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        NumSerText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosC(NumDocText.Text) = "" Then
        MsgBox "No ha especificado un nùmero de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        NumDocText.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosC(InventarioTextBox.Text) = "" Then
        MsgBox "No ha especificado un nombre para el inventario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        InventarioTextBox.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosN(IdResponsableText.Text) = 0 Then
        MsgBox "No ha especificado un responsable para el inventario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        InventarioTextBox.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosN(IdAlmacenTextBox.Text) = 0 Then
        MsgBox "No ha especificado un nombre para el almacén", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        IdAlmacenTextBox.SetFocus
        validarDatos = False
        Exit Function
    End If
    If NulosN(IdTipoInventarioTextBox.Text) = 0 Then
        MsgBox "No ha especificado un tipo de inventario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        IdTipoInventarioTextBox.SetFocus
        validarDatos = False
        Exit Function
    End If
    If fg(0).Rows = fg(0).FixedRows Then
        MsgBox "No han especificado registros para grabar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        validarDatos = False
        Exit Function
    End If
    
    validarDatos = True
End Function

Function Grabar(Optional MostrarMensajes As Boolean = True) As Boolean
    Dim Rpta As Integer
    Dim A As Integer
    Dim Preguntar As Boolean
    Dim Inventario As New AlmacenEntidad.EInventario
    Dim ObjInv As New AlmacenEntidad.EInventario
        
On Error GoTo LaCague
    If Not validarDatos Then Grabar = False: Exit Function
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "Grabar", "Modificar") + " el registro de inventario", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function
    
    If QueHace = 2 Then
        If (NulosN(IdTipoInventarioTextBox.Text) = F.NuloNumeric(F.KeyValue("InventarioAjuste", xCon))) Then
            If (NulosN(RstPro("idestadoinventario")) = F.NuloNumeric(F.KeyValue("EstadoAprobadoInventario", xCon))) Then
                If MsgBox("El documento actual ya se encuentra registrado, ¿Desea realizar un reajuste?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
                    ObjInv.IdInventario = NulosN(RstPro("idtomainventario"))
                    Set ObjInv.Conexion = xCon
                    If ObjInv.DeleteMovimientos(0, "") Then
                        cargaStockCosto
                    Else
                        Err.Raise &HFFFFFF01, , "Error al intentar eliminar los movimientos del ajuste"
                    End If
                End If
            End If
        End If
    End If
            
    ' Cabecera
    If QueHace = 1 Then Inventario.IdInventario = 0 Else Inventario.IdInventario = NulosN(RstPro("idtomainventario"))
    Inventario.IdAlmacen = NulosN(IdAlmacenTextBox.Text)
    Inventario.IdTipoInventario = NulosN(IdTipoInventarioTextBox.Text)
    Inventario.IdEstado = NulosN(IdEstadoLabel.Caption)
    Inventario.NumeroSerie = NulosC(NumSerText.Text)
    Inventario.NumeroDocumento = NulosC(NumDocText.Text)
    Inventario.Descripcion = NulosC(InventarioTextBox.Text)
    Inventario.IdResponsable = NulosN(IdResponsableText.Text)
    Inventario.Glosa = NulosC(DescripcionTextBox.Text)
    Inventario.FechaInventario = FchInvTextBoxFecha.Valor
    Inventario.FechaVigencia = FchVigTextBoxFecha.Valor
    If OptTipo(0).Value = True Then ' Todos
        Inventario.IdTipoFiltro = 0
    ElseIf OptTipo(1).Value = True Then ' Con movimiento
        Inventario.IdTipoFiltro = 1
    End If
    Inventario.AnhoTrabajo = AnoTra
    
    ' Detalle
    For A = fg(0).FixedRows To fg(0).Rows - 1
        Dim InvDetalle As New AlmacenEntidad.EInventarioDet
        
        InvDetalle.IdItem = NulosN(fg(0).TextMatrix(A, fg(0).ColIndex("IDITEM")))
        InvDetalle.IdUnidadMedida = NulosN(fg(0).TextMatrix(A, fg(0).ColIndex("IDUNIMED")))
        InvDetalle.CantidadInicial = NulosN(fg(0).TextMatrix(A, fg(0).ColIndex("STOCKACT")))
        InvDetalle.CostoInicial = NulosN(fg(0).TextMatrix(A, fg(0).ColIndex("PREUNIACT")))
        InvDetalle.CantidadCarga = NulosN(fg(0).TextMatrix(A, fg(0).ColIndex("CANTIDAD")))
        InvDetalle.CostoCarga = NulosN(fg(0).TextMatrix(A, fg(0).ColIndex("PREUNI")))
        Inventario.InventarioDetS.Add InvDetalle
        Set InvDetalle = Nothing
    Next A
    
    Set Inventario.Conexion = xCon
    If Not Inventario.Save(0, "") Then Err.Raise &HFFFFFF01, , "Error al intentar grabar el registro"
    
    MsgBox "El inventario se grabó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    mIdRegistro = Inventario.IdInventario
    Set Inventario = Nothing
    Grabar = True
    Exit Function
    
LaCague:
    Set Inventario = Nothing
    Grabar = False
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description), vbCritical, xTitulo
End Function

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELAR EL PROCESO DE AGRGAR O MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    Bloquea
    fg(0).Editable = flexEDNone
    fg(0).SelectionMode = flexSelectionByRow
    ActivaTool
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    frm(4).Visible = False
    frm(0).Visible = False
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA UN REGISTRO PARA SU MODIFICACION
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    QueHace = 2
    Bloquea
    ActivaTool
    If TabOne1.CurrTab = 0 Then
        Blanquea
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If

    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Modificando Inventario"
    
    fg(0).Editable = flexEDKbdMouse
    fg(0).SelectionMode = flexSelectionFree
    xHorIni = Time
    IdAlmacenTextBox.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA EL REGISTRO ACTUAL
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    Dim Inventario As New AlmacenEntidad.EInventario
    
On Error GoTo BloqueError
    TabOne1.CurrTab = 0
    If RstPro.State = 0 Then Exit Sub
    If RstPro.RecordCount = 0 Then
        MsgBox "No hay registros para eliminar", vbInformation, xTitulo
        Exit Sub
    End If
    
    ' SI EL ITEM NO TIENE NINGUNA OPERACION SE PROCEDE A ELIMINAR PREVIA AUTORIZACION DEL USUARIO
    Rpta = MsgBox("¿ Esta seguro de eliminar el registro ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        Inventario.IdInventario = NulosN(RstPro("idtomainventario"))
        Inventario.IdEstado = NulosN(RstPro("idestadoinventario"))
        Inventario.IdTipoInventario = NulosN(RstPro("idtipoinventario"))
        Set Inventario.Conexion = xCon
        If Inventario.Delete(0, "") Then
            MsgBox "El registro se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            RstPro.Requery
            Dg1.Refresh
        Else
            Err.Raise &HFFFFFF01, , "Error al intentar eliminar el registro"
        End If
        Exit Sub
    End If
    Exit Sub
    
BloqueError:
    MsgBox "No se pudo eliminar el registro por el siguiente motivo :" + Trim(Err.Description)
    Set Inventario = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DEL REGISTRO ACTUAL, ESTE EVENTO SE EJECUTA CUANDO EL
'*                    FORMULARIO ESTA EN MODO DE LECTURA O MODIFICAR
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    Dim xRs As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer
    Dim F As New SistemaLogica.Funciones
    
    If RstPro.RecordCount = 0 Then Exit Sub
    If RstPro.BOF = True Or RstPro.EOF = True Then Exit Sub
    
    '************
    ' Cabecera
    '************
    IdInventarioLabel.Caption = NulosN(RstPro("idtomainventario"))
    FchInvTextBoxFecha.Valor = RstPro("fchinv")
    FchVigTextBoxFecha.Valor = RstPro("fchvig")
    NumSerText.Text = NulosC(RstPro("numser"))
    NumDocText.Text = NulosC(RstPro("numdoc"))
    IdEstadoLabel.Caption = NulosN(RstPro("idestadoinventario"))
    EstadoLabel.Caption = NulosC(RstPro("estadoinventario"))
    IdAlmacenTextBox.Text = NulosN(RstPro("idalm"))
    AlmacenLabel.Caption = NulosC(RstPro("almacen"))
    IdTipoInventarioTextBox.Text = NulosN(RstPro("idtipoinventario"))
    TipoInventarioLabel.Caption = NulosC(RstPro("tipoinventario"))
    InventarioTextBox.Text = NulosC(RstPro("nombre"))
    IdResponsableText.Text = NulosN(RstPro("idresponsable"))
    ResponsableLabel.Caption = NulosC(RstPro("responsable"))
    DescripcionTextBox.Text = NulosC(RstPro("descripcion"))
    OptTipo(NulosN(RstPro("tipofiltro"))).Value = True
    
    '************
    ' Detalle
    '************
    cSQL = "SELECT alm_tomainventariodet.iditem, alm_tomainventariodet.idunimed, alm_tomainventariodet.stockactual, alm_tomainventariodet.preuniactual, alm_tomainventariodet.cantidad, alm_tomainventariodet.preuni, alm_inventario.codpro, alm_inventario.descripcion As item, mae_unidades.abrev As unimed " _
        + vbCr + "FROM (alm_tomainventariodet INNER JOIN alm_inventario ON alm_tomainventariodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON alm_tomainventariodet.idunimed = mae_unidades.id " _
        + vbCr + "WHERE (((alm_tomainventariodet.idtomainventario) = " & NulosN(RstPro("idtomainventario")) & "));"
    
    Set RstDet = Nothing
    RST_Busq RstDet, cSQL, xCon
    
    Agregando = True
    pConfigurarGrid fg(0)
    fg(0).Rows = fg(0).FixedRows
    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        For A = 1 To RstDet.RecordCount
            fg(0).Rows = fg(0).Rows + 1
            fg(0).TextMatrix(A + 1, fg(0).ColIndex("IDITEM")) = NulosN(RstDet("iditem"))
            fg(0).TextMatrix(A + 1, fg(0).ColIndex("IDUNIMED")) = NulosN(RstDet("idunimed"))
            fg(0).TextMatrix(A + 1, fg(0).ColIndex("STOCKACT")) = Format(NulosN(RstDet("stockactual")), FORMAT_CANTIDAD)
            fg(0).TextMatrix(A + 1, fg(0).ColIndex("PREUNIACT")) = Format(NulosN(RstDet("preuniactual")), FORMAT_MONTO)
            fg(0).TextMatrix(A + 1, fg(0).ColIndex("CANTIDAD")) = Format(NulosN(RstDet("cantidad")), FORMAT_CANTIDAD)
            fg(0).TextMatrix(A + 1, fg(0).ColIndex("PREUNI")) = Format(NulosN(RstDet("preuni")), FORMAT_MONTO)
            fg(0).TextMatrix(A + 1, fg(0).ColIndex("DIFCANTIDAD")) = Format(NulosN(RstDet("cantidad")) - NulosN(RstDet("stockactual")), FORMAT_CANTIDAD)
            fg(0).TextMatrix(A + 1, fg(0).ColIndex("DIFPREUNI")) = Format(NulosN(RstDet("preuni")) - NulosN(RstDet("preuniactual")), FORMAT_CANTIDAD)
            fg(0).TextMatrix(A + 1, fg(0).ColIndex("CODIGO")) = NulosC(RstDet("codpro"))
            fg(0).TextMatrix(A + 1, fg(0).ColIndex("ITEM")) = NulosC(RstDet("item"))
            fg(0).TextMatrix(A + 1, fg(0).ColIndex("UM")) = NulosC(RstDet("unimed"))
            RstDet.MoveNext
            
            If RstDet.EOF = True Then
                Exit For
            End If
        Next A
    End If
    
    ' Total a Producir
    fg(0).Select fg(0).FixedRows, fg(0).ColIndex("DIFCANTIDAD"), fg(0).Rows - 1, fg(0).ColIndex("DIFCANTIDAD")
    fg(0).FillStyle = flexFillRepeat
    fg(0).CellBackColor = &HC0FFFF       ' &H00000000&
    
    ' Total a Producir
    fg(0).Select fg(0).FixedRows, fg(0).ColIndex("DIFPREUNI"), fg(0).Rows - 1, fg(0).ColIndex("DIFPREUNI")
    fg(0).FillStyle = flexFillRepeat
    fg(0).CellBackColor = &HC0FFFF
    
    lblnItm.Caption = fg(0).Rows - fg(0).FixedRows
    F.PintarGrid fg(0), fg(0).ColIndex("DIFCANTIDAD"), &H0&, &HFF&
    F.PintarGrid fg(0), fg(0).ColIndex("DIFPREUNI"), &H0&, &HFF&
    Agregando = False
    bloquearControles
End Sub

Sub ExportarExcel(ByRef GRID_ As VSFlexGrid)
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    Dim TITULO_ As String
    
    TITULO_ = "Formato de Inventario"
    
    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, GRID_, TITULO_, "", "", TITULO_
    Set X_EXPORT = Nothing
    MousePointer = vbDefault
    Exit Sub
error:
    MousePointer = vbDefault
    SHOW_ERROR Name, "Exportar"
End Sub

Private Function BuscaItemExcelGrilla(mIdItem As Long) As Integer
    Dim A As Integer
    
    ' Valida Items Grilla
    If fg(0).Rows = fg(0).FixedRows Then
        BuscaItemExcelGrilla = 0
        Exit Function
    End If
    
    ' Valida en la Grilla
    For A = fg(0).FixedRows To fg(0).Rows - 1
        If F.NuloString(fg(0).TextMatrix(A, fg(0).ColIndex("IDITEM"))) = mIdItem Then
            BuscaItemExcelGrilla = A
            Exit Function
        End If
    Next
    
    BuscaItemExcelGrilla = 0
End Function

Private Sub ImportarExcel()
    'Especificar las extensiones a usar
    Dim nPath As String
    Dim mRowAdd As Double
    Cmm.DefaultExt = "*.xls"
    Cmm.Filter = "Documentos de Excel (*.xls)|*.xls"
    Cmm.ShowOpen
    If Err Then
        'Cancelada la operación de abrir
        GoTo error
    Else
        nPath = Cmm.FileName
    End If
    
    If nPath = "" Then GoTo error

    Dim A&
    Dim mNumeroFilas As Integer
    Dim mFilaInicial As Integer
    Dim mFilaFinal As Integer
    Dim mFilaActual As Integer
    Dim rstEmp As New ADODB.Recordset
    Dim RstTar As New ADODB.Recordset
    Dim RstProd As New ADODB.Recordset
    Dim nSQL As String
    Dim IdItem As Long
    Dim mFilaActualGrilla As Long
    Dim IdUniMed As Integer
    Dim objExcel As Object
    
    Me.MousePointer = vbHourglass
    
    Set objExcel = CreateObject("Excel.Application")
    'objExcel.Visible = True
    objExcel.SheetsInNewWorkbook = 1
    
    'objExcel.WindowState = 2
    objExcel.Workbooks.Open nPath
    
    CentrarFrm frm(5)
    lblProcesado.Caption = "Cargando registros para la importación"
    frm(5).Visible = True
    
    mNumeroFilas = 0
    mFilaInicial = 10
    mFilaActual = mFilaInicial
    ' Se determina la fila final
    While Not NulosC(objExcel.ActiveSheet.Cells(mFilaActual, 2)) = ""
        mNumeroFilas = mNumeroFilas + 1
        mFilaActual = mFilaActual + 1
    Wend
    mFilaFinal = mFilaInicial + mNumeroFilas
    
    lblProcesado.Caption = "Importando Inventario"
    lbl(32).Caption = "Importando:"
        
    Agregando = True
    PgBar.Max = mFilaFinal
    PgBar.Min = mFilaInicial
    PgBar.Value = mFilaInicial
    
    For A = mFilaInicial To mFilaFinal
        PgBar.Value = A
        mFilaActual = A
        IdItem = 0
        IdUniMed = 0
        ' Se valida si existe o no el item
        IdItem = NulosN(Busca_Codigo(NulosC(objExcel.ActiveSheet.Cells(mFilaActual, 2)), "codpro", "id", "alm_inventario", "C", xCon))
        If IdItem > 0 Then
            mFilaActualGrilla = BuscaItemExcelGrilla(IdItem)
            ' Si no se encuentra en la grilla
            If mFilaActualGrilla = 0 Then
                fg(0).Rows = fg(0).Rows + 1
                mFilaActualGrilla = fg(0).Rows - 1
                ' Se cargan los datos
                fg(0).TextMatrix(mFilaActualGrilla, fg(0).ColIndex("CODIGO")) = NulosC(Busca_Codigo(IdItem, "id", "codpro", "alm_inventario", "N", xCon))
                fg(0).TextMatrix(mFilaActualGrilla, fg(0).ColIndex("ITEM")) = NulosC(Busca_Codigo(IdItem, "id", "descripcion", "alm_inventario", "N", xCon))
                fg(0).TextMatrix(mFilaActualGrilla, fg(0).ColIndex("IDITEM")) = IdItem
                ' Se valida si existe o no la unidad de medida
                IdUniMed = NulosN(Busca_Codigo(NulosC(objExcel.ActiveSheet.Cells(mFilaActual, 4)), "abrev", "id", "mae_unidades", "C", xCon))
                If IdUniMed = 0 Then
                    ' Se busca la unidad de medida por defecto para el item
                    IdUniMed = NulosN(Busca_Codigo(IdItem, "id", "idunimed", "alm_inventario", "N", xCon))
                End If
                fg(0).TextMatrix(mFilaActualGrilla, fg(0).ColIndex("IDUNIMED")) = IdUniMed
                fg(0).TextMatrix(mFilaActualGrilla, fg(0).ColIndex("UM")) = Busca_Codigo(IdUniMed, "id", "abrev", "mae_unidades", "N", xCon)
                ' Se cargan valores actuales de cantidades
                If (NulosN(IdTipoInventarioTextBox.Text) = NulosN(F.KeyValue("InventarioInicial", xCon))) Then
                    fg(0).TextMatrix(mFilaActualGrilla, fg(0).ColIndex("STOCKACT")) = Format(NulosN(Busca_Codigo(IdItem, "id", "stckini", "alm_inventario", "N", xCon)), FORMAT_CANTIDAD)
                    fg(0).TextMatrix(mFilaActualGrilla, fg(0).ColIndex("PREUNIACT")) = Format(F.CostoInicial(IdItem, F.NuloNumeric(IdAlmacenTextBox.Text), xCon), FORMAT_MONTO)
                    
                ElseIf (NulosN(IdTipoInventarioTextBox.Text) = NulosN(F.KeyValue("InventarioAjuste", xCon))) Then
                    fg(0).TextMatrix(mFilaActualGrilla, fg(0).ColIndex("STOCKACT")) = Format(F.SaldoActual(IdItem, F.NuloNumeric(IdAlmacenTextBox.Text), "01/01/" & AnoTra, FchInvTextBoxFecha.Valor, xCon), FORMAT_CANTIDAD)
                    fg(0).TextMatrix(mFilaActualGrilla, fg(0).ColIndex("PREUNIACT")) = Format(F.CostoActual(IdItem, F.NuloNumeric(IdAlmacenTextBox.Text), "01/01/" & AnoTra, FchInvTextBoxFecha.Valor, xCon), FORMAT_MONTO)
                End If
            End If
            ' Se muestra en pantalla
            lblProcesado.Caption = NulosC(fg(0).TextMatrix(mFilaActualGrilla, fg(0).ColIndex("ITEM")))
            frm(5).Refresh
            ' Se cargan valores del excel
            fg(0).TextMatrix(mFilaActualGrilla, fg(0).ColIndex("CANTIDAD")) = Format(NulosN(objExcel.ActiveSheet.Cells(mFilaActual, 6)), FORMAT_CANTIDAD)
            fg(0).TextMatrix(mFilaActualGrilla, fg(0).ColIndex("PREUNI")) = Format(F.NuloNumeric(objExcel.ActiveSheet.Cells(mFilaActual, 9)), FORMAT_MONTO)
            fg(0).TextMatrix(mFilaActualGrilla, fg(0).ColIndex("DIFCANTIDAD")) = Format(NulosN(fg(0).TextMatrix(mFilaActualGrilla, fg(0).ColIndex("CANTIDAD"))) - NulosN(fg(0).TextMatrix(mFilaActualGrilla, fg(0).ColIndex("STOCKACT"))), FORMAT_CANTIDAD)
            fg(0).TextMatrix(mFilaActualGrilla, fg(0).ColIndex("DIFPREUNI")) = Format(NulosN(fg(0).TextMatrix(mFilaActualGrilla, fg(0).ColIndex("PREUNI"))) - NulosN(fg(0).TextMatrix(mFilaActualGrilla, fg(0).ColIndex("PREUNIACT"))), FORMAT_MONTO)
        End If
    Next A
    
    frm(5).Visible = False
    Agregando = False
    
    MsgBox "El proceso terminó de cargar los datos con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    objExcel.WindowState = 2
    objExcel.Workbooks.Close
    
    Set objExcel = Nothing
    Me.MousePointer = vbDefault
    lblnItm.Caption = fg(0).Rows - fg(0).FixedRows
    Exit Sub

error:
    Agregando = False
    frm(5).Visible = False
    Me.MousePointer = vbDefault
    Set objExcel = Nothing
    SHOW_ERROR Me.Name, "pImportar"
End Sub

Private Sub CargarExcel()
    'Especificar las extensiones a usar
    Dim nPath As String
    Dim mRowAdd As Double
    Cmm.DefaultExt = "*.xls"
    Cmm.Filter = "Documentos de Excel (*.xls)|*.xls"
    Cmm.ShowOpen
    If Err Then
        'Cancelada la operación de abrir
        GoTo error
    Else
        nPath = Cmm.FileName
    End If
    
    If nPath = "" Then GoTo error

    Dim A&
    Dim mNumeroFilas As Integer
    Dim mFilaInicial As Integer
    Dim mFilaActual As Integer
    Dim rstEmp As New ADODB.Recordset
    Dim RstTar As New ADODB.Recordset
    Dim RstProd As New ADODB.Recordset
    Dim nSQL As String
    Dim objExcel As Object
    
    Me.MousePointer = vbHourglass
    
    Set objExcel = CreateObject("Excel.Application")
    'objExcel.Visible = True
    objExcel.SheetsInNewWorkbook = 1
    
    'objExcel.WindowState = 2
    objExcel.Workbooks.Open nPath
    
    CentrarFrm frm(5)
    lblProcesado.Caption = "Cargando registros para la importación"
    frm(5).Visible = True
    
    mNumeroFilas = 1
    mFilaInicial = 10
    mFilaActual = mFilaInicial
    ' DETERMINAMOS EL NUMERO DE FILAS CON DATOS
    While Not NulosC(objExcel.ActiveSheet.Cells(mFilaActual, 2)) = ""
        mNumeroFilas = mNumeroFilas + 1
        mFilaActual = mFilaActual + 1
    Wend
    
    lblProcesado.Caption = "Importando Inventario"
    lbl(32).Caption = "Importando:"
        
    Agregando = True
    PgBar.Max = fg(0).Rows - 1
    PgBar.Value = 0
    For A = fg(0).FixedRows To fg(0).Rows - 1
        PgBar.Value = A
        mFilaActual = mFilaInicial
        lblProcesado.Caption = NulosC(fg(0).TextMatrix(A, fg(0).ColIndex("ITEM")))
        frm(5).Refresh
        
        ' Se busca al mismo nivel de la grilla
        If NulosC(objExcel.ActiveSheet.Cells(mFilaInicial + (A - fg(0).FixedRows), 2)) = NulosC(fg(0).TextMatrix(A, fg(0).ColIndex("CODIGO"))) Then
            mFilaActual = mFilaInicial + (A - fg(0).FixedRows)
        Else
            ' Se busca el registro en el Excel
            Do While mFilaActual <= mNumeroFilas + mFilaInicial
                If NulosC(objExcel.ActiveSheet.Cells(mFilaActual, 2)) = NulosC(fg(0).TextMatrix(A, fg(0).ColIndex("CODIGO"))) Then
                    Exit Do
                Else
                    mFilaActual = mFilaActual + 1
                End If
            Loop
        End If
        ' Se cargan los datos
        fg(0).TextMatrix(A, fg(0).ColIndex("CANTIDAD")) = Format(NulosN(objExcel.ActiveSheet.Cells(mFilaActual, 6)), FORMAT_CANTIDAD)
        fg(0).TextMatrix(A, fg(0).ColIndex("PREUNI")) = Format(F.NuloNumeric(objExcel.ActiveSheet.Cells(mFilaActual, 9)), FORMAT_MONTO)
        fg(0).TextMatrix(A, fg(0).ColIndex("DIFCANTIDAD")) = Format(NulosN(fg(0).TextMatrix(A, fg(0).ColIndex("CANTIDAD"))) - NulosN(fg(0).TextMatrix(A, fg(0).ColIndex("STOCKACT"))), FORMAT_CANTIDAD)
        fg(0).TextMatrix(A, fg(0).ColIndex("DIFPREUNI")) = Format(NulosN(fg(0).TextMatrix(A, fg(0).ColIndex("PREUNI"))) - NulosN(fg(0).TextMatrix(A, fg(0).ColIndex("PREUNIACT"))), FORMAT_MONTO)
    Next A
    
    frm(5).Visible = False
    Agregando = False
    
    MsgBox "El proceso termino de cargar los datos con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    objExcel.WindowState = 2
    objExcel.Workbooks.Close
    
    Set objExcel = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Agregando = False
    frm(5).Visible = False
    Me.MousePointer = vbDefault
    Set objExcel = Nothing
    SHOW_ERROR Me.Name, "pImportar"
End Sub

'Metodos para arrastrar el Frame
''''''''''''''''''''''''''''''''
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

