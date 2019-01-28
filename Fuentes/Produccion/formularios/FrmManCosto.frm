VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManCosto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Producción - Configurar Costo"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9855
   Begin VB.Frame FraHoras 
      BorderStyle     =   0  'None
      Height          =   4230
      Left            =   7410
      TabIndex        =   46
      Top             =   2760
      Visible         =   0   'False
      Width           =   9420
      Begin VB.CommandButton CmdTarea 
         Caption         =   "Buscar"
         Height          =   330
         Index           =   3
         Left            =   4050
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   330
         Width           =   735
      End
      Begin VB.TextBox TxtCosto3 
         Height          =   300
         Left            =   5670
         MaxLength       =   12
         TabIndex        =   57
         Text            =   "TxtCosto3"
         Top             =   405
         Width           =   510
      End
      Begin VB.CommandButton CmdTarea 
         Caption         =   "Calcular"
         Height          =   330
         Index           =   4
         Left            =   6180
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   375
         Width           =   735
      End
      Begin VB.CommandButton CmdTarea 
         Caption         =   "Grabar"
         Height          =   330
         Index           =   5
         Left            =   7440
         TabIndex        =   49
         Top             =   360
         Width           =   945
      End
      Begin VB.CommandButton CmdTarea 
         Caption         =   "Cancelar"
         Height          =   330
         Index           =   6
         Left            =   8370
         TabIndex        =   48
         Top             =   360
         Width           =   945
      End
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   9120
         Picture         =   "FrmManCosto.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   47
         ToolTipText     =   "Cerrar"
         Top             =   30
         Width           =   195
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
         Height          =   300
         Index           =   0
         Left            =   705
         TabIndex        =   51
         Top             =   360
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
      Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
         Height          =   300
         Index           =   1
         Left            =   2670
         TabIndex        =   52
         Top             =   360
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
      Begin VSFlex7Ctl.VSFlexGrid fg1 
         Height          =   3225
         Index           =   2
         Left            =   90
         TabIndex        =   55
         Top             =   840
         Width           =   9225
         _cx             =   16272
         _cy             =   5689
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
         Rows            =   1
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmManCosto.frx":02EC
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora2"
         Height          =   195
         Left            =   5190
         TabIndex        =   58
         Top             =   510
         Width           =   435
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   2115
         TabIndex        =   54
         Top             =   465
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   150
         TabIndex        =   53
         Top             =   465
         Width           =   510
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   15
         Y1              =   15
         Y2              =   5070
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   0
         X1              =   30
         X2              =   9360
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   0
         X2              =   9300
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   9390
         X2              =   9390
         Y1              =   -120
         Y2              =   4770
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Actualizar Costo"
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
         Left            =   75
         TabIndex        =   50
         Top             =   60
         Width           =   1395
      End
      Begin VB.Line Line3 
         X1              =   60
         X2              =   9330
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   270
         Index           =   0
         Left            =   30
         Top             =   15
         Width           =   9315
      End
   End
   Begin VB.Frame FraCosto 
      BorderStyle     =   0  'None
      Height          =   4230
      Left            =   8880
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   9420
      Begin VB.CommandButton CmdTarea 
         Caption         =   "Calcular"
         Height          =   330
         Index           =   2
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox TxtCosto2 
         Height          =   300
         Left            =   570
         MaxLength       =   12
         TabIndex        =   43
         Text            =   "TxtCosto2"
         Top             =   390
         Width           =   510
      End
      Begin VB.CommandButton cb1 
         Height          =   240
         Index           =   0
         Left            =   3195
         Picture         =   "FrmManCosto.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   420
         Width           =   225
      End
      Begin SizerOneLibCtl.TabOne TabOne2 
         Height          =   3255
         Left            =   120
         TabIndex        =   33
         Top             =   870
         Width           =   9225
         _cx             =   16272
         _cy             =   5741
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
         Caption         =   " Directo  | Diverso "
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
            BorderStyle     =   0  'None
            Height          =   3225
            Left            =   10155
            TabIndex        =   36
            Top             =   15
            Width           =   8880
            Begin VSFlex7Ctl.VSFlexGrid fg1 
               Height          =   3165
               Index           =   1
               Left            =   60
               TabIndex        =   42
               Top             =   60
               Width           =   8760
               _cx             =   15452
               _cy             =   5583
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
               Rows            =   1
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManCosto.frx":0586
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
         Begin VB.Frame Frame5 
            BorderStyle     =   0  'None
            Height          =   3225
            Left            =   330
            TabIndex        =   34
            Top             =   15
            Width           =   8880
            Begin VSFlex7Ctl.VSFlexGrid fg1 
               Height          =   3135
               Index           =   0
               Left            =   60
               TabIndex        =   35
               Top             =   60
               Width           =   8760
               _cx             =   15452
               _cy             =   5530
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
               Rows            =   1
               Cols            =   12
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManCosto.frx":0681
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
      Begin VB.TextBox TxtCosto1 
         Height          =   300
         Left            =   1380
         MaxLength       =   12
         TabIndex        =   30
         Text            =   "TxtCosto1"
         Top             =   240
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   9120
         Picture         =   "FrmManCosto.frx":07CD
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   29
         ToolTipText     =   "Cerrar"
         Top             =   30
         Width           =   195
      End
      Begin VB.CommandButton CmdTarea 
         Caption         =   "Cancelar"
         Height          =   330
         Index           =   1
         Left            =   8370
         TabIndex        =   28
         Top             =   360
         Width           =   945
      End
      Begin VB.CommandButton CmdTarea 
         Caption         =   "Aceptar"
         Height          =   330
         Index           =   0
         Left            =   7440
         TabIndex        =   27
         Top             =   360
         Width           =   945
      End
      Begin VB.TextBox txt_cb1 
         Height          =   300
         Index           =   0
         Left            =   2655
         MaxLength       =   12
         TabIndex        =   38
         Text            =   "txt_cb1(0)"
         Top             =   390
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora2"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lbl_cod1 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cod1(0)"
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
         Height          =   270
         Index           =   0
         Left            =   4560
         TabIndex        =   40
         Top             =   420
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label lbl_cb_capt1 
         AutoSize        =   -1  'True
         Caption         =   "Tarea"
         Height          =   195
         Index           =   1
         Left            =   2130
         TabIndex        =   39
         Top             =   480
         Width           =   420
      End
      Begin VB.Line Line9 
         X1              =   60
         X2              =   9330
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Actualizar Costo"
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
         Left            =   75
         TabIndex        =   32
         Top             =   60
         Width           =   1395
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   9390
         X2              =   9390
         Y1              =   -120
         Y2              =   4770
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -150
         X2              =   9360
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   2
         X1              =   30
         X2              =   9360
         Y1              =   4200
         Y2              =   4200
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
      Begin VB.Label lblCosto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora1"
         Height          =   195
         Left            =   930
         TabIndex        =   31
         Top             =   330
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   270
         Index           =   2
         Left            =   30
         Top             =   15
         Width           =   9315
      End
      Begin VB.Label lbl_cb1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cb1(0)"
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
         Left            =   3480
         TabIndex        =   41
         Top             =   390
         Width           =   3870
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
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCosto.frx":0AB9
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCosto.frx":0FFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCosto.frx":138F
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCosto.frx":1513
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCosto.frx":1967
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCosto.frx":1A7F
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCosto.frx":1FC3
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCosto.frx":2507
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCosto.frx":261B
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCosto.frx":272F
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCosto.frx":2B83
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCosto.frx":2CEF
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCosto.frx":3237
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCosto.frx":35C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManCosto.frx":38E3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   0
      TabIndex        =   4
      Top             =   375
      Width           =   9855
      _cx             =   17383
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
         Caption         =   "Detalle de la Cuenta"
         Height          =   6795
         Left            =   10500
         TabIndex        =   8
         Top             =   375
         Width           =   9765
         Begin VSFlex7Ctl.VSFlexGrid fg 
            Height          =   1710
            Index           =   1
            Left            =   75
            TabIndex        =   22
            Top             =   3705
            Width           =   9615
            _cx             =   16960
            _cy             =   3016
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
            Rows            =   10
            Cols            =   14
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmManCosto.frx":3BFD
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
         Begin VB.Frame Frame9 
            Height          =   645
            Left            =   75
            TabIndex        =   19
            Top             =   5445
            Width           =   9630
            Begin VB.CommandButton Command1 
               Caption         =   "Ver Histórico"
               Height          =   375
               Left            =   7620
               TabIndex        =   25
               Top             =   150
               Visible         =   0   'False
               Width           =   1890
            End
            Begin VB.CommandButton cmd 
               Caption         =   "Agregar"
               Enabled         =   0   'False
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   21
               ToolTipText     =   "Agregar Tarea"
               Top             =   180
               Width           =   1590
            End
            Begin VB.CommandButton cmd 
               Caption         =   "&Eliminar"
               Enabled         =   0   'False
               Height          =   375
               Index           =   1
               Left            =   2175
               TabIndex        =   20
               TabStop         =   0   'False
               ToolTipText     =   "Eliminar Tarea"
               Top             =   180
               Width           =   1590
            End
         End
         Begin VB.TextBox txt 
            BackColor       =   &H0080FF80&
            Height          =   315
            Index           =   0
            Left            =   8520
            TabIndex        =   17
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   195
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Frame FraReceta 
            Caption         =   "[ Receta ]"
            Enabled         =   0   'False
            Height          =   675
            Left            =   2955
            TabIndex        =   12
            Top             =   420
            Width           =   6720
            Begin VB.CommandButton cb 
               Height          =   240
               Index           =   0
               Left            =   1980
               Picture         =   "FrmManCosto.frx":3D96
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   285
               Width           =   225
            End
            Begin VB.TextBox txt_cb 
               Height          =   300
               Index           =   0
               Left            =   1020
               MaxLength       =   12
               TabIndex        =   2
               Text            =   "txt_cb(0)"
               Top             =   255
               Width           =   1215
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
               Left            =   3435
               TabIndex        =   15
               Top             =   255
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.Label lbl_cb_capt 
               AutoSize        =   -1  'True
               Caption         =   "Descripción"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   14
               Top             =   360
               Width           =   840
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
               Left            =   2220
               TabIndex        =   16
               Top             =   255
               Width           =   4320
            End
         End
         Begin VB.Frame FraTipo 
            Caption         =   "[ Origen ]"
            Enabled         =   0   'False
            Height          =   675
            Left            =   90
            TabIndex        =   11
            Top             =   420
            Width           =   2835
            Begin VB.OptionButton opt_origen 
               Caption         =   "Tarea Diversa"
               Height          =   195
               Index           =   1
               Left            =   1275
               TabIndex        =   1
               Top             =   315
               Width           =   1350
            End
            Begin VB.OptionButton opt_origen 
               Caption         =   "Receta"
               Height          =   195
               Index           =   0
               Left            =   180
               TabIndex        =   0
               Top             =   300
               Value           =   -1  'True
               Width           =   990
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid fg 
            Height          =   1950
            Index           =   0
            Left            =   90
            TabIndex        =   3
            Top             =   1350
            Width           =   5595
            _cx             =   9869
            _cy             =   3440
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
            Rows            =   10
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmManCosto.frx":3EC8
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   75
            TabIndex        =   24
            Top             =   3435
            Width           =   6330
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tareas"
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
            Left            =   120
            TabIndex        =   23
            Top             =   1125
            Width           =   600
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "codigo"
            Height          =   195
            Index           =   0
            Left            =   8010
            TabIndex        =   18
            Top             =   225
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Costo"
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
            Left            =   60
            TabIndex        =   9
            Top             =   30
            Width           =   9660
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6795
         Left            =   45
         TabIndex        =   5
         Top             =   375
         Width           =   9765
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   30
            TabIndex        =   6
            Top             =   345
            Width           =   9690
            _ExtentX        =   17092
            _ExtentY        =   11324
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Id"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Origen"
            Columns(1).DataField=   "origen"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Descripción"
            Columns(2).DataField=   "referencia"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Receta"
            Columns(3).DataField=   "codrec"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2461"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2381"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=9208"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=9128"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=2487"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2408"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(52)  =   "Named:id=33:Normal"
            _StyleDefs(53)  =   ":id=33,.parent=0"
            _StyleDefs(54)  =   "Named:id=34:Heading"
            _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(56)  =   ":id=34,.wraptext=-1"
            _StyleDefs(57)  =   "Named:id=35:Footing"
            _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(59)  =   "Named:id=36:Selected"
            _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(61)  =   "Named:id=37:Caption"
            _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(63)  =   "Named:id=38:HighlightRow"
            _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(65)  =   "Named:id=39:EvenRow"
            _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(67)  =   "Named:id=40:OddRow"
            _StyleDefs(68)  =   ":id=40,.parent=33"
            _StyleDefs(69)  =   "Named:id=41:RecordSelector"
            _StyleDefs(70)  =   ":id=41,.parent=34"
            _StyleDefs(71)  =   "Named:id=42:FilterBar"
            _StyleDefs(72)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Costo"
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
            Left            =   30
            TabIndex        =   7
            Top             =   30
            Width           =   9690
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
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
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Actualizar Costo en Lote"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Equivalencia de Costo en Horas"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmManCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMMANCOSTO.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO QUE PERMITE ASIGNAR COSTO A LAS TAREAS A CADA RECETA
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 05/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstFrm As New ADODB.Recordset       ' RECORDSET QUE ALAMCENARA LOS PRODCUTOS DISPONIBLES
Dim QueHace As Integer                  ' ESPECIFICA EN QUE MODO SE ENCUENTRA EL FORMULARIO
Dim Agregando As Boolean                ' INDICA QUE SE ESTAN AGREGANDO FILAS A UN CONTROL FLEXGRID
Dim SeEjecuto As Boolean                ' CONTROLA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim mIdRegistro&                        ' identificador del registro
Dim fOrdenLista As Boolean              ' especfica el orden de la lista de la consulta
Dim RstValores As New ADODB.Recordset   '
Dim mRowAdd As Double                   ' identificador unico por fila cuando se agrege una unidad
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim xHorIni As Date  'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO

'*****************************************************************************************************
'* Nombre           : Filtrar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA UN FILTRO SOBRE EL RECORDSET RstFrm
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Filtrar()
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(3, 4) As String
   
    xCampos(0, 0) = "Código":           xCampos(0, 1) = "codigo":        xCampos(0, 2) = "C":         xCampos(0, 3) = "4200"
    xCampos(1, 0) = "Descripcion":      xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "C":         xCampos(1, 3) = "1500"
    xCampos(2, 0) = "Abrev":            xCampos(2, 1) = "abrev":         xCampos(2, 2) = "C":         xCampos(2, 3) = "200"
    TabOne1.CurrTab = 0
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg1
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0 ' agregar
            pRegistroAdd
            
        Case 1 ' eliminar
            pRegistroDel
    End Select
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstFrm
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA EN FORMA ASCENDENTE O DECENDENTE LAS COLUMNAS DEL CONTROL Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstFrm("id")), xCon
    End If
End Sub

Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If QueHace = 3 Then Exit Sub
    If Index = 0 Then Exit Sub
    
    Select Case Col
        Case 1, 2
            Exit Sub
            
        Case 3, 4, 5, 6, 7, 8, 9, 11, 12, 13
            If Col <> 12 And Col <> 13 Then
                If NulosN(fg(1).TextMatrix(Row, Col)) <> 0 Then
                    If IsNumeric(fg(1).TextMatrix(Row, Col)) = False Then
                        MsgBox "El valor ingresado no es numérico", vbCritical, xTitulo
                        fg(1).TextMatrix(Row, Col) = ""
                    End If
                End If
            Else
                If Trim(fg(1).TextMatrix(Row, Col)) = "/  /" Then
                    fg(1).TextMatrix(Row, Col) = ""
                ElseIf IsDate(fg(1).TextMatrix(Row, Col)) = False Then
                    MsgBox "El valor ingresado no es una fecha Correcta", vbCritical, xTitulo
                    fg(1).TextMatrix(Row, Col) = ""
                End If
            End If
            
            ' aplicando filtro
            RstValores.Filter = "idtar = " & fg(0).TextMatrix(fg(0).Row, 2)
            If RstValores.RecordCount <> 0 Then RstValores.MoveFirst
            RstValores.Find "idunimed = " & NulosN(fg(1).TextMatrix(Row, 10))
            
            If RstValores.EOF = False And RstValores.BOF = False Then
            
                RstValores("minimo") = NulosN(fg(1).TextMatrix(Row, 3))
                RstValores("promedio") = NulosN(fg(1).TextMatrix(Row, 4))
                RstValores("maximo") = NulosN(fg(1).TextMatrix(Row, 5))
                RstValores("cant") = NulosN(fg(1).TextMatrix(Row, 6))
                RstValores("jornal") = NulosN(fg(1).TextMatrix(Row, 7))
                RstValores("costo") = NulosN(fg(1).TextMatrix(Row, 8))
                RstValores("orden") = NulosN(fg(1).TextMatrix(Row, 9))
                RstValores("canreg") = NulosN(fg(1).TextMatrix(Row, 11))
                If IsDate((fg(1).TextMatrix(Row, 12))) = True Then RstValores("fchini") = CDate((fg(1).TextMatrix(Row, 12)))
                If IsDate((fg(1).TextMatrix(Row, 13))) = True Then RstValores("fchfin") = CDate((fg(1).TextMatrix(Row, 13)))
                If IsDate((fg(1).TextMatrix(Row, 12))) = False Then RstValores("fchini") = Null
                If IsDate((fg(1).TextMatrix(Row, 13))) = False Then RstValores("fchfin") = Null
            End If
    End Select
End Sub

Private Sub Fg_EnterCell(Index As Integer)
    If QueHace = 3 Then
        fg(Index).Editable = flexEDNone
        Exit Sub
    End If
    
    If Index = 0 Then
        If opt_origen(0).Value = True And fg(Index).Col = 1 Then
            fg(Index).Editable = flexEDNone
        Else
            fg(Index).Editable = flexEDKbdMouse
        End If
    Else
        If fg(Index).Col <> 1 Then
            fg(Index).Editable = flexEDKbdMouse
        Else
            fg(Index).Editable = flexEDNone
        End If
    End If
End Sub

Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If QueHace = 3 Then Exit Sub
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLNotIn As String
    Dim nSQLTmp  As String
    Dim nTitulo As String
    
    If Index = 0 Then
        If Col <> 1 Then Exit Sub
        ' de la tarea
        ReDim xCampos(3, 4) As String
        xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":     xCampos(0, 2) = "4500":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "Abrev":        xCampos(1, 1) = "nomcorto":   xCampos(1, 2) = "2300":    xCampos(1, 3) = "C"
        xCampos(2, 0) = "Id":           xCampos(2, 1) = "id":         xCampos(2, 2) = "600":     xCampos(2, 3) = "N"
        
        If NulosC(fg(Index).TextMatrix(Row, Col)) <> "" Then
            nSQLTmp = " AND UCASE(pro_tareas.descripcion) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' "
        End If
        
        nSQL = "SELECT pro_tareas.id, pro_tareas.codigo, pro_tareas.descripcion AS nombre, pro_tareas.abrev AS nomcorto, mae_unidades.id AS idunimed, mae_unidades.abrev " _
            + vbCr + " FROM mae_unidades RIGHT JOIN pro_tareas ON mae_unidades.id = pro_tareas.idunimed " _
            + vbCr + " WHERE (((pro_tareas.id) Not In (SELECT pro_costo.idref FROM pro_costo WHERE pro_costo.tipo=2)) )  " & nSQLTmp
        nTitulo = "Buscando Tareas"
    Else
        Select Case Col
            Case 2
                ' de la unidad de medida
                ReDim xCampos(2, 4) As String
                xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":   xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
                xCampos(1, 0) = "Abrev":        xCampos(1, 1) = "abrev":    xCampos(1, 2) = "800":    xCampos(1, 3) = "C"
                xCampos(2, 0) = "Id":           xCampos(2, 1) = "id":       xCampos(2, 2) = "600":    xCampos(2, 3) = "N"
                nSQL = "SELECT mae_unidades.id, mae_unidades.descripcion as nombre, mae_unidades.abrev FROM mae_unidades ;"
                
            Case 12, 13
                ' invocar al formulario de fecha
                Dim obj As New SGI2_funciones.formularios
                obj.FechaSeleccionar fg(1), Row, Col, fg(1).TextMatrix(Row, Col), e_Seleccion
                Set obj = Nothing
                Exit Sub
            
            Case Else
                Exit Sub
                
        End Select
    End If
    
    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio, ""

    If RstTmp.State = 0 Then GoTo SALIR
    If RstTmp.EOF = True Or RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR
    
    Agregando = True
    If Index = 0 Then
        ' si la tarea es diferente a la seleccionada => eliminar lista de unidades
        If NulosN(fg(0).TextMatrix(fg(0).Row, 2)) <> NulosN(RstTmp("id")) Then
            RstValores.Filter = ""
            RstRegistroEliminar RstValores, "idtar", fg(0).TextMatrix(fg(0).Row, 2), True
            ' inicializar el grid
            fg(1).Rows = 1
        End If
        ' llenando los datos
        fg(Index).TextMatrix(Row, 1) = NulosC(RstTmp("nombre"))
        fg(Index).TextMatrix(Row, 2) = NulosC(RstTmp("id"))
    Else
        ' aplicando filtro
        RstValores.Filter = "idtar = " & fg(0).TextMatrix(fg(0).Row, 2)
        If RstValores.RecordCount <> 0 Then RstValores.MoveFirst
        RstValores.Find "idunimed =" & NulosN(fg(Index).TextMatrix(Row, 10))
        If RstValores.EOF = True Or RstValores.BOF = True Then RstValores.AddNew
        
        RstValores("idtar") = fg(0).TextMatrix(fg(0).Row, 2)
        RstValores("idunimed") = NulosN(RstTmp("id"))
        RstValores("descripcion") = NulosC(RstTmp("nombre"))
        RstValores("abrev") = NulosC(RstTmp("abrev"))
        
        ' llenando los datos
        fg(Index).TextMatrix(Row, 1) = NulosC(RstTmp("nombre"))
        fg(Index).TextMatrix(Row, 2) = NulosC(RstTmp("abrev"))
        fg(Index).TextMatrix(Row, 10) = NulosC(RstTmp("id"))
    End If
    
    Agregando = False
    Set RstTmp = Nothing
    Exit Sub

SALIR:
    Set RstTmp = Nothing
    Agregando = False
End Sub

Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If Col = 2 Then
        If validar_letras(KeyAscii) = False Then
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub Fg_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo error
    If Index = 0 Then Exit Sub
    If KeyCode = 114 Or KeyCode = vbKeyInsert Then cmd_Click 0 'F3 = Agregar Item
    If KeyCode = 115 Or KeyCode = vbKeyDelete Then cmd_Click 1         'F4 = Eliminar Item
    Exit Sub

error:
    SHOW_ERROR Me.Name, "Fg_KeyUp (" & Index & ")"
End Sub

Private Sub fg_RowColChange(Index As Integer)
    If Agregando = True Then Exit Sub
    
    If Index = 1 Then Exit Sub
    
    If fg(0).Row < 1 Then
        habilitar Cmd, False
        Exit Sub
    End If
    
    Label2.Caption = fg(0).TextMatrix(fg(0).Row, 1)
    
    If QueHace <> 3 Then habilitar Cmd, True
    If RstValores.State = 0 Then Exit Sub
    
    ' Mostramos los insumos de la receta
    RstValores.Filter = adFilterNone
    RstValores.Filter = "idtar = " & NulosN(fg(0).TextMatrix(fg(0).Row, 2))
    If RstValores.RecordCount <> 0 Then
        pCargarDatosValores
    Else
        fg(1).Rows = 1
    End If
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    On Error GoTo error
    If SeEjecuto = False Then
        
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        RST_Busq RstFrm, "SELECT pro_costo.id, pro_costo.tipo, pro_costo.idref, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas].[descripcion]) AS referencia, IIf([pro_costo].[tipo]=1,[pro_receta].[codrec],'') AS codrec " _
            + vbCr + " FROM ((alm_inventario RIGHT JOIN pro_receta ON alm_inventario.id = pro_receta.iditem) RIGHT JOIN pro_costo ON pro_receta.id = pro_costo.idref) LEFT JOIN pro_tareas ON pro_costo.idref = pro_tareas.id " _
            + vbCr + " GROUP BY pro_costo.id, pro_costo.tipo, pro_costo.idref, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa'), IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas].[descripcion]), pro_receta.codrec " _
            + vbCr + " ORDER BY pro_costo.tipo, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa'),IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas].[descripcion])", xCon

        Set Dg1.DataSource = RstFrm

        
        ' ocultando las columnas de codigos
        OCULTAR_COL Fg1(0), 7, 11
        OCULTAR_COL Fg1(1), 5, 8
        
        OCULTAR_COL Fg1(2), 7, 11
        OCULTAR_COL Fg1(2), 3, 3 ' receta
        
        Fg1(0).ColFormat(5) = "###,##0.00"
        Fg1(0).ColFormat(6) = "#0.000000"
        
        Fg1(1).ColFormat(3) = "###,##0.00"
        Fg1(1).ColFormat(4) = "#0.000000"
        
        Fg1(2).ColFormat(5) = "###,##0.00"
        Fg1(2).ColFormat(6) = "#0.000000"
        TxtFecha(0).valor = Date
        TxtFecha(1).valor = Date
    End If
    Exit Sub

error:
    SHOW_ERROR Me.Name, "Form_Activate"
End Sub

'*****************************************************************************************************
'* Nombre           : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Nuevo()
    Bloquea False
    Blanquea
    ActivaTool
    QueHace = 1
    xHorIni = Time
    opt_origen(0).Value = True
    Label5.Caption = "Agregando Costo"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS BOTONES DE LA BARRA DE HERRAMIENTAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA LOS CONTROLES TEXTBOX PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    LimpiaText txt
    LimpiaText txt_cb
    LimpiaText lbl_cod
    LimpiaText txt_cb
    fg(0).Rows = 1
    fg(1).Rows = 1
End Sub

'*****************************************************************************************************
'* Nombre           : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES TEXTBOX Y COMMAND
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Bloquea(band As Boolean)
    habilitar_Locked txt, band
    habilitar_Locked txt_cb, band
    habilitar Cmd, Not band
    FraTipo.Enabled = Not band
    FraReceta.Enabled = Not band
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    CentrarFrm Me
    QueHace = 3
    TabOne1.CurrTab = 0
    
    ' ocultando las columnas
    fg(0).ColWidth(2) = 0   ' idtarea
    fg(1).ColWidth(10) = 0  ' idunidad
    
    GRID_COMBOLIST fg(0), 1
    GRID_COMBOLIST fg(1), 2 '--UM
    
    fg(1).ColFormat(3) = FORMAT_MONTO
    fg(1).ColFormat(4) = FORMAT_MONTO
    fg(1).ColFormat(5) = FORMAT_MONTO
    fg(1).ColFormat(6) = FORMAT_MONTO
    
    fg(1).ColFormat(7) = "0.0000000"
    fg(1).ColFormat(8) = "0.0000000"
    
    fg(1).ColEditMask(12) = "##/##/####" ' fecha inicio
    fg(1).ColEditMask(13) = "##/##/####" ' fecha fin
    GRID_COMBOLIST fg(1), 12
    GRID_COMBOLIST fg(1), 13
    
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO EN LA PESTAÑA DETALLE DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    Blanquea
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then Exit Sub

    If RstFrm("tipo") = 1 Then
        opt_origen(0).Value = True
    Else
        opt_origen(1).Value = True
    End If
    
    txt_cb(0).Text = NulosC(RstFrm("codrec"))
    lbl_cb(0).Caption = NulosC(RstFrm("referencia"))
    lbl_cod(0).Caption = NulosN(RstFrm("idref"))
    
    MuestraDetalle
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL PROCESO DE AGREGAR O MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    Bloquea True
    ActivaTool
    Label5.Caption = "Detalle del Costo"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Dg1.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario mientras este ingresando o modificando un Costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    Else
        Set RstFrm = Nothing
        SeEjecuto = False
    End If
End Sub

Private Sub opt_origen_Click(Index As Integer)
    fg(0).Rows = 1
    fg(1).Rows = 1
    If Index = 0 Then
        FraReceta.Visible = True
    Else
        FraReceta.Visible = False
        txt_cb(0).Text = "***":     txt_cb(0).Text = ""
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 3 Then MuestraSegundoTab
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
            
            RstFrm.MoveFirst
            RstFrm.Find "id = " & mIdRegistro & ""
            If RstFrm.EOF = True Then
                RstFrm.MoveFirst
            End If
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 8 Then Filtrar
    
    If Button.Index = 9 Then
        If RstFrm.State = 0 Then Exit Sub
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstFrm.Filter = adFilterNone
        RstFrm.Requery
        Dg1.Refresh
    End If
    
    If Button.Index = 10 Then Buscar
    
    If Button.Index = 12 Then pExportar
    
    If Button.Index = 15 Then pHabilitarBotonEditor True, 1
    
    If Button.Index = 16 Then pHabilitarBotonEditor True, 2
    
    If Button.Index = 18 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : FUNCION
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA pro_costodet, ESTA FUNCION DEVUELVE VERDADERO
'*                    CUANDO TIENE EXITO
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Function Grabar() As Boolean
    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " el Costo", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo SALIR
       
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xId As Double
    Dim xCol&, xFil&, xCorr&
    
    
    On Error GoTo LaCague

    xCon.BeginTrans
    Me.MousePointer = vbHourglass
    
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT top 1 * FROM pro_costo ", xCon
        xId = HallaCodigoTabla("pro_costo", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstFrm("id")
        RST_Busq RstCab, "SELECT * FROM pro_costo WHERE id =" & xId & "", xCon
        xCon.Execute "DELETE * FROM pro_costodet WHERE idcos = " & xId & ""
    End If
    
    mIdRegistro = xId
    
    RST_Busq RstDet, "SELECT top 1 * FROM pro_costodet", xCon
    
    ' 1 receta; 2::tarea diversa
    If opt_origen(0).Value = True Then
        RstCab("tipo") = 1
        RstCab("idref") = NulosN(lbl_cod(0).Caption)
    Else
        RstCab("tipo") = 2
        RstCab("idref") = NulosN(fg(0).TextMatrix(1, 2))
    End If
    
    RstCab.Update
    ' recorrer las tareas
    For xFil = 1 To fg(0).Rows - 1
        ' registro de las unidades
        xCorr = 1
        RstValores.Filter = "idtar= " & NulosN(fg(0).TextMatrix(xFil, 2))
        If RstValores.RecordCount <> 0 Then RstValores.MoveFirst
        
        Do While Not RstValores.EOF
            RstDet.AddNew
            ' codigo
            RstDet("idcos") = xId
            RstDet("idtar") = NulosN(fg(0).TextMatrix(xFil, 2))
            RstDet("corr") = xCorr
            ' fin codigo
            RstDet("idunimed") = NulosN(RstValores.Fields("idunimed"))
            RstDet("minimo") = NulosN(RstValores.Fields("minimo"))
            RstDet("promedio") = NulosN(RstValores.Fields("promedio"))
            RstDet("maximo") = NulosN(RstValores.Fields("maximo"))
            RstDet("cant") = NulosN(RstValores.Fields("cant"))
            RstDet("jornal") = NulosN(RstValores.Fields("jornal"))
            RstDet("costo") = NulosN(RstValores.Fields("costo"))
            RstDet("orden") = NulosN(RstValores.Fields("orden"))
            
            If IsDate(RstValores.Fields("fchini")) = True Then RstDet("fchini") = CDate(RstValores.Fields("fchini"))
            If IsDate(RstValores.Fields("fchfin")) = True Then RstDet("fchfin") = CDate(RstValores.Fields("fchfin"))
            
            xCorr = xCorr + 1
            RstDet.Update
            RstValores.MoveNext
        Loop
    Next xFil
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    
    xCon.CommitTrans
    MsgBox "El Costo se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    Grabar = True

SALIR:
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing
    Exit Function

LaCague:
    xCon.RollbackTrans
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
    Grabar = False
End Function

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA LA EDICION DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    Bloquea False
    Blanquea
    ActivaTool
    QueHace = 2
    xHorIni = Time
    Label5.Caption = "Modificando Costo"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    MuestraSegundoTab
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA pro_costo
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    Dim RstTmp  As New ADODB.Recordset
    Dim nSQL As String
    Dim xId&
    
    TabOne1.CurrTab = 0
    
    xId = NulosN(RstFrm.Fields("id"))

    Set RstTmp = Nothing
    Rpta = MsgBox("Esta seguro de eliminar el Costo seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM pro_costodet WHERE idcos =" & xId & " "
        xCon.Execute "DELETE * FROM pro_costo WHERE id =" & xId & " "
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & xId & " AND idform = " & IdMenuActivo
        
        MsgBox "El Costo se eliminó con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstFrm.Requery
        Dg1.Refresh
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA UNA BUSQUEDA EN EL RECORDSET RstFrm
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Buscar()
    On Error GoTo error
    Dim RstTmp As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Tipo":           xCampos(0, 1) = "Origen":       xCampos(0, 2) = "1200":        xCampos(0, 3) = "c"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "Referencia":   xCampos(1, 2) = "4500":        xCampos(1, 3) = "c"
    xCampos(2, 0) = "Cod.Rec":        xCampos(2, 1) = "codrec":       xCampos(2, 2) = "900":         xCampos(2, 3) = "c"
    
    TabOne1.CurrTab = 0
    
    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, RstFrm.Source, xCampos(), "Buscando Costo", "Referencia", "Referencia", Principio
    If RstTmp.State = 0 Then GoTo SALIR
    If RstTmp.EOF = True And RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR
    RstFrm.MoveFirst
    RstFrm.Find "id = " & RstTmp("id") & ""

SALIR:
    Set RstTmp = Nothing

error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub

'*****************************************************************************************************
'* Nombre           : fValidarDatos
'* Tipo             : FUNCCION
'* Descripcion      : VERIFICA QUE LOS DATOS INGRESADO SEAN LOS CORRECTOS, ESTA FUNCION DEVUELVE
'*                    VERDADERO
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Private Function fValidarDatos() As Boolean
    If opt_origen(0).Value = True And NulosN(lbl_cod(0).Caption) = 0 Then
        MsgBox "Falta especificar la Receta", vbExclamation, xTitulo
        txt_cb(0).SetFocus
        Exit Function
    End If
    
    If fg(0).Rows = 0 Then
        MsgBox "Falta Especificar las Tarea", vbExclamation, xTitulo
        Exit Function
    End If
    
    fValidarDatos = True
End Function
 
Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cod(Index).Caption = ""
        If Index = 0 Then
            Label2.Caption = ""
            ' limpiar los datos
            Agregando = True
            Set RstValores = Nothing
            If opt_origen(0).Value = True Then
                fg(0).Rows = 1
            Else
                fg(0).Rows = 2
                fg(0).TextMatrix(1, 1) = ""
                fg(0).TextMatrix(1, 2) = ""
            End If
            fg(1).Rows = 1
            ' cargar otra vez la data
            pCargarDatosRstTemp -10
            Agregando = False
        End If
    End If
End Sub

Private Sub txt_cb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    
    If txt_cb(Index).Locked = True Then Exit Sub
    
    If KeyCode = vbKeyF5 Then
        cb_Click Index
        Exit Sub
    End If
End Sub

Private Sub txt_cb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
End Sub

Private Sub cb_Click(Index As Integer)
    ' EJECUTA LA BUSQUEDA DE RECETAS
    If QueHace = 3 Then Exit Sub
    Dim RstTmp As New ADODB.Recordset
    Dim xCampos() As String
    Dim nSQL As String
    On Error GoTo error

    ReDim xCampos(3, 4) As String
    xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "codpro":   xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "nombre":   xCampos(1, 2) = "5000":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "CodReceta":    xCampos(2, 1) = "codrec":   xCampos(2, 2) = "1200":    xCampos(2, 3) = "C"
       
    nSQL = "SELECT DISTINCT pro_receta.codrec, alm_inventario.descripcion AS nombre, pro_receta.id AS cod,alm_inventario.codpro, pro_receta.iditem " _
        + vbCr + " FROM alm_inventario INNER JOIN pro_receta ON alm_inventario.id = pro_receta.iditem " _
        + vbCr + " WHERE (((pro_receta.id) In (SELECT pro_recetatar.idrec FROM pro_recetatar ) And (pro_receta.id) Not In (SELECT pro_costo.idref FROM pro_costo WHERE pro_costo.tipo = 1))) " _
        + vbCr + " ORDER BY alm_inventario.descripcion; "
    
    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), "Buscando Recetas", "nombre", "nombre", Principio
    
    If RstTmp.State = 0 Then GoTo SALIR
    If RstTmp.EOF = True Or RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR
    
    If NulosN(lbl_cod(Index).Caption) <> NulosN(RstTmp.Fields(2)) Then
        txt_cb(Index).Text = NulosC(RstTmp.Fields(0))     ' TEXTO A MOSTRAR
        lbl_cb(Index).Caption = NulosC(RstTmp.Fields(1))  ' NOMBRE
        lbl_cod(Index).Caption = NulosN(RstTmp.Fields(2)) ' CODIGO
    
        nSQL = "SELECT pro_tareas.descripcion, pro_recetatar.idtar " _
            + vbCr + " FROM pro_tareas INNER JOIN pro_recetatar ON pro_tareas.id = pro_recetatar.idtar " _
            + vbCr + " WHERE (((pro_recetatar.idrec) = " & NulosN(RstTmp.Fields(2)) & " )) " _
            + vbCr + " ORDER BY pro_recetatar.orden;"
        
        Set RstTmp = Nothing
        
        RST_Busq RstTmp, nSQL, xCon
        fg(0).Rows = 1
        Agregando = True
        If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            DoEvents
            fg(0).Rows = fg(0).Rows + 1
            fg(0).TextMatrix(fg(0).Rows - 1, 1) = NulosC(RstTmp("descripcion"))
            fg(0).TextMatrix(fg(0).Rows - 1, 2) = NulosN(RstTmp("idtar"))
            RstTmp.MoveNext
        Loop
        
        Agregando = False
        ' cargar el recordset
        pCargarDatosRstTemp -10
        
        If fg(0).Rows > 1 Then
            fg(0).Row = 1:  fg(0).Col = 1:  fg(0).SetFocus
        Else
            txt_cb(0).SetFocus
        End If
    End If
        
SALIR:
    Set RstTmp = Nothing
    Exit Sub

error:
    Agregando = False
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

'*****************************************************************************************************
'* Nombre           : pExportar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA A MS EXCEL LOS DATOS DEL RECORDSET RSTTMP
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pExportar()
    TabOne1.CurrTab = 0
        
    Dim xCampos(3, 3) As String
    Dim oExport As New SGI2_funciones.formularios
    Dim RstTmp As New ADODB.Recordset
    Set RstTmp = RstFrm.Clone
    ' 0 Nombre a Mostrar;
    ' 1 nombre de Campo del Rst;
    ' 2 alineacion(0::derecha, 1::centro, 2::izquierda);
    ' 3 ancho de columna
    ' obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Código":       xCampos(0, 1) = "codigo":       xCampos(0, 2) = 0:  xCampos(0, 3) = "1200"
    xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = 0:  xCampos(1, 3) = "3500"
    xCampos(2, 0) = "Unidad":       xCampos(2, 1) = "abrev":        xCampos(2, 2) = 0:  xCampos(2, 3) = "750"
    xCampos(3, 0) = "Es Diverso":   xCampos(3, 1) = "diverso":      xCampos(3, 2) = 0:  xCampos(3, 3) = "800"
    
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "Lista de Costos", "", "", "Listado de Costo", RstTmp, xCampos()
    
    Set oExport = Nothing
    Set RstTmp = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraDetalle
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO ACTUAL
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub MuestraDetalle()
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    ' limpiando el rst temporal
    Set RstValores = Nothing
    
    nSQL = "SELECT distinct  IIf([pro_costo].[tipo]=1,[pro_recetatar].[idtar],[pro_tareas].[id]) AS idtar, IIf([pro_costo].[tipo]=1,[pro_tareas_1].[descripcion],[pro_tareas].[descripcion]) AS tarea " _
        + vbCr + " FROM ((pro_recetatar LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_recetatar.idtar = pro_tareas_1.id) RIGHT JOIN pro_costo ON pro_recetatar.idrec = pro_costo.idref) LEFT JOIN pro_tareas ON pro_costo.idref = pro_tareas.id " _
        + vbCr + " WHERE (((pro_costo.id) = " & RstFrm("id") & ")) " _
        + vbCr + "  "

    RST_Busq RstTmp, nSQL, xCon
    
    DoEvents
    If RstTmp.RecordCount <> 0 Then
        DoEvents
        Agregando = True
        With fg(0)
            .Rows = 1
            RstTmp.MoveFirst
            Do While Not RstTmp.EOF
                DoEvents
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = NulosC(RstTmp.Fields("tarea"))
                .TextMatrix(.Rows - 1, 2) = NulosN(RstTmp.Fields("idtar"))
                RstTmp.MoveNext
            Loop
        End With
    End If
    
    Set RstTmp = Nothing
    
    ' cargar datos de los valores de las tareas
    pCargarDatosRstTemp NulosN(RstFrm("id"))
    
    If fg(0).Rows > 1 Then
        fg(0).Row = 1:  fg(0).Col = 1:
        Agregando = False
        fg_RowColChange 0
    End If
    Agregando = False
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Set RstTmp = Nothing
    Agregando = False
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "MuestraDetalle"
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarDatosRstTemp
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : Definir la estructura del recordset de los valores, ESTA FUNCION DEVUELVE UN
'*                    RECORDSET CON DATOS
'* Paranetros       : NOMBRE    |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    idCodigo  |  INTEGER    |  codigo del Costo
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarDatosRstTemp(idCodigo)
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Set RstTmp = Nothing
    
    ' definir la estructura de recordset
    nSQL = "SELECT mae_unidades.descripcion, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant, pro_costodet.costo, pro_costodet.orden, pro_costodet.idtar, pro_costodet.idunimed,pro_costodet.fchini, pro_costodet.fchfin, pro_costodet.canreg " _
            + vbCr + " FROM mae_unidades RIGHT JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed " _
            + vbCr + " WHERE (((pro_costodet.idcos)=" & idCodigo & " )) " _
            + vbCr + " ORDER BY mae_unidades.descripcion; "
            
    RST_Busq RstTmp, nSQL, xCon
    DEFINIR_RST_TMP RstValores, RstTmp
    
    Set RstTmp = Nothing
    RST_Busq RstTmp, nSQL, xCon
    DoEvents
    DEFINIR_RST_TMP RstValores, RstTmp
    If RstTmp.RecordCount <> 0 Then CARGAR_RST_TMP RstValores, RstTmp
    Set RstTmp = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarDatosValores
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL DETALLE DEL REGISTRO ACTUAL
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarDatosValores()
    fg(1).Rows = 1
    Agregando = True
    If RstValores.RecordCount <> 0 Then
        RstValores.MoveFirst
        RstValores.Sort = "fchini"
        With fg(1)
            Do While Not RstValores.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = NulosC(RstValores("descripcion"))
                .TextMatrix(.Rows - 1, 2) = NulosC(RstValores("abrev"))
                .TextMatrix(.Rows - 1, 3) = NulosN(RstValores("minimo"))
                .TextMatrix(.Rows - 1, 4) = NulosN(RstValores("promedio"))
                .TextMatrix(.Rows - 1, 5) = NulosN(RstValores("maximo"))
                .TextMatrix(.Rows - 1, 6) = NulosN(RstValores("cant"))
                .TextMatrix(.Rows - 1, 7) = NulosN(RstValores("jornal"))
                .TextMatrix(.Rows - 1, 8) = NulosN(RstValores("costo"))
                .TextMatrix(.Rows - 1, 9) = NulosN(RstValores("orden"))
                .TextMatrix(.Rows - 1, 10) = NulosN(RstValores("idunimed"))
                .TextMatrix(.Rows - 1, 11) = NulosN(RstValores("canreg"))
                .TextMatrix(.Rows - 1, 12) = Format(NulosC(RstValores("fchini")), FORMAT_DATE)
                .TextMatrix(.Rows - 1, 13) = Format(NulosC(RstValores("fchfin")), FORMAT_DATE)
                RstValores.MoveNext
            Loop
        End With
    End If
    Agregando = False
End Sub

'*****************************************************************************************************
'* Nombre           : pRegistroAdd
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : AÑADE UNA FILA AL CONTROL Fg
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pRegistroAdd()
    Dim mCol%
    Dim fInsertar As Boolean
    Agregando = True
    If fg(0).Row < 1 Then
        MsgBox "Falta especificar la Tarea", vbInformation, xTitulo
        Exit Sub
    End If
    If NulosN(fg(0).TextMatrix(fg(0).Row, 2)) = 0 Then
        MsgBox "Falta especificar la Tarea", vbInformation, xTitulo
        Exit Sub
    End If
    If fg(1).Rows > fg(1).FixedRows Then
        If NulosC(fg(1).TextMatrix(fg(1).Rows - 1, 1)) = "" Then    ' descripcion de unidad
            MsgBox "Seleccione la Unidad", vbInformation, xTitulo
        Else
            fInsertar = True
        End If
    Else
        fInsertar = True
    End If
    mCol = 2
    
    If fInsertar = True Then fg(1).AddItem ""
    fg(1).Row = fg(1).Rows - 1
    fg(1).Col = mCol
    If fInsertar = True Then Fg_CellButtonClick 1, fg(1).Rows - 1, 1
    fg(1).SetFocus
    Agregando = False
End Sub

'*****************************************************************************************************
'* Nombre           : pRegistroDel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UNA FILA DEL CONTROL Fg
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pRegistroDel()
    If fg(1).Row < 1 Then
        MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(1).SetFocus
        Exit Sub
    End If
    
    If fg(1).Rows = 1 Then
        MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(1).SetFocus
        Exit Sub
    End If
    If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    
    ' aplicando filtro
    RstValores.Filter = "idtar = " & fg(0).TextMatrix(fg(0).Row, 2)
    If RstValores.RecordCount <> 0 Then RstValores.MoveFirst
    RstValores.Find "idunimed  = " & NulosN(fg(1).TextMatrix(fg(1).Row, 10))
    If RstValores.EOF = False And RstValores.BOF = False Then
        RstValores.Delete
    End If
    fg(1).RemoveItem fg(1).Row
 
    If fg(1).Rows > 1 Then
        fg(1).Row = fg(1).Rows - 1
        fg(1).Col = 1
        fg(1).SetFocus
    Else
        Cmd(0).SetFocus
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : pHabilitarBotonEditor
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : Actualizar el costo, Costo Actualizado en funcion al costo por hora
'* Paranetros       : NOMBRE    |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    band      |  Boolean    |  puede ser true o false
'*                    tipo      |  Integer    |  tipo= 1(actualizar horas en lote), 2(Convertir el pago
'*                                               a horas)
'* Devuelve         :
'*****************************************************************************************************
Private Sub pHabilitarBotonEditor(band As Boolean, Tipo As Integer)
    ' true muestra el ingreso de datos
    If band = True Then
        TabOne1.CurrTab = 0
        If Tipo = 1 Then
            FraCosto.Top = 1530
            FraCosto.Left = 30
            TxtCosto2.Text = ""
            txt_cb1(0).Text = ""
            lbl_cb1(0).Caption = ""
        Else
            FraHoras.Top = 1950
            FraHoras.Left = 30
            TxtCosto3.Text = ""
            Fg1(2).Rows = 1
        End If
    End If
    
    If Tipo = 1 Then
        TxtCosto1.Enabled = band
        TxtCosto2.Enabled = band
        txt_cb1(0).Locked = Not band
        cb1(0).Enabled = band
        FraCosto.Visible = band
    Else
        TxtCosto3.Enabled = band
        FraHoras.Visible = band
    End If
    
    Toolbar1.Enabled = Not band
    TabOne1.Enabled = Not band
    
    If band = True And Tipo = 1 Then pDatosCostoCargar
    ' si es true cargar los datos
End Sub

Private Sub CmdTarea_Click(Index As Integer)
    Select Case Index
        Case 0 ' aceptar
            pDatosCostoGrabar
        
        Case 1 ' cancelar
            pHabilitarBotonEditor False, 1
        
        Case 2 ' calcular
            pDatosCostoCalcular 1
      
        Case 3 ' buscar tareas
            pConvertAHoraCargar
        
        Case 4 ' calcular igual que case 2
            pDatosCostoCalcular 2
        
        Case 5 ' aceptar
            pConvertAHoraGrabar
        
        Case 6 ' cancelar
            pHabilitarBotonEditor False, 2
    End Select
    
End Sub

Private Sub Fg1_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    
    If Col = 12 Then Exit Sub
    
    If Index <> 2 Then
        If NulosN(TxtCosto2.Text) = 0 Then
            MsgBox "Ingrese el Costo por Destino", vbExclamation, xTitulo
            TxtCosto2.SetFocus
            Exit Sub
        End If
    Else
        If NulosN(TxtCosto3.Text) = 0 Then
            MsgBox "Ingrese el Costo por Destino", vbExclamation, xTitulo
            TxtCosto3.SetFocus
            Exit Sub
        End If
    End If
    Agregando = True
    
    Select Case Index
        Case 0 ' tareas relacionado con producto
            If Col = 5 Then
                If IsNumeric(Fg1(0).TextMatrix(Row, Col)) = False Then
                    Fg1(0).TextMatrix(Row, Col) = 0
                End If
                If NulosN(Fg1(0).TextMatrix(Row, Col)) <> 0 Then
                    Fg1(0).TextMatrix(Row, 6) = NulosN(TxtCosto2.Text) / NulosN(Fg1(0).TextMatrix(Row, Col))
                Else
                    Fg1(0).TextMatrix(Row, 6) = 0
                End If
            ElseIf Col = 6 Then
                If NulosN(Fg1(0).TextMatrix(Row, Col)) <> 0 Then
                    Fg1(0).TextMatrix(Row, 5) = NulosN(TxtCosto2.Text) / NulosN(Fg1(0).TextMatrix(Row, Col))
                Else
                    Fg1(0).TextMatrix(Row, 5) = 0
                End If
               
            End If
            ' aplicando color
            If NulosN(Fg1(0).TextMatrix(Row, 6)) = 0 Then
                GRID_COLOR_FONDO Fg1(0), Row, 6, Row, 6, vbRed
            Else
                GRID_COLOR_FONDO Fg1(0), Row, 6, Row, 6, vbWhite
            End If
            Fg1(0).Row = Row
            Fg1(0).Col = Col
        
        Case 1 ' tareas diversas
            If Col = 3 Then
                If IsNumeric(Fg1(1).TextMatrix(Row, Col)) = False Then
                    Fg1(1).TextMatrix(Row, Col) = 0
                End If
                If NulosN(Fg1(1).TextMatrix(Row, Col)) <> 0 Then
                    Fg1(1).TextMatrix(Row, 4) = NulosN(TxtCosto2.Text) / NulosN(Fg1(1).TextMatrix(Row, Col))
                Else
                    Fg1(1).TextMatrix(Row, 4) = 0
                End If
            ElseIf Col = 4 Then
                If NulosN(Fg1(1).TextMatrix(Row, Col)) <> 0 Then
                    Fg1(1).TextMatrix(Row, 3) = NulosN(TxtCosto2.Text) / NulosN(Fg1(1).TextMatrix(Row, Col))
                Else
                    Fg1(1).TextMatrix(Row, 3) = 0
                End If
            End If
            ' aplicando color
            If NulosN(Fg1(1).TextMatrix(Row, 4)) = 0 Then
                GRID_COLOR_FONDO Fg1(1), Row, 4, Row, 4, vbRed
            Else
                GRID_COLOR_FONDO Fg1(1), Row, 4, Row, 4, vbWhite
            End If
            Fg1(1).Row = Row
            Fg1(1).Col = Col
        
        Case 2
            If Col = 5 Then
                If IsNumeric(Fg1(2).TextMatrix(Row, Col)) = False Then
                    Fg1(2).TextMatrix(Row, Col) = 0
                End If
                If NulosN(Fg1(2).TextMatrix(Row, Col)) <> 0 Then
                    Fg1(2).TextMatrix(Row, 6) = NulosN(TxtCosto3.Text) / NulosN(Fg1(2).TextMatrix(Row, Col))
                Else
                    Fg1(2).TextMatrix(Row, 6) = 0
                End If
            ElseIf Col = 6 Then
                If NulosN(Fg1(2).TextMatrix(Row, Col)) <> 0 Then
                    Fg1(2).TextMatrix(Row, 5) = NulosN(TxtCosto3.Text) / NulosN(Fg1(2).TextMatrix(Row, Col))
                Else
                    Fg1(2).TextMatrix(Row, 5) = 0
                End If
               
            End If
            ' aplicando color
            If NulosN(Fg1(2).TextMatrix(Row, 6)) = 0 Then
                GRID_COLOR_FONDO Fg1(2), Row, 6, Row, 6, vbRed
            Else
                GRID_COLOR_FONDO Fg1(2), Row, 6, Row, 6, vbWhite
            End If
            Fg1(2).Row = Row
            Fg1(2).Col = Col
    End Select
    
    Agregando = False
End Sub

Private Sub Fg1_EnterCell(Index As Integer)
    If Agregando = True Then Exit Sub
    If Index = 0 Then
        If Fg1(0).Col = 5 Or Fg1(0).Col = 6 Then
            If NulosN(TxtCosto2.Text) = 0 Then
                MsgBox "Ingrese el costo por Hora", vbExclamation, xTitulo
                TxtCosto2.SetFocus
                Fg1(0).Editable = flexEDNone
                Exit Sub
            End If
            Fg1(0).Editable = flexEDKbdMouse
        Else
            Fg1(0).Editable = flexEDNone
        End If
    ElseIf Index = 1 Then
        If Fg1(1).Col = 3 Or Fg1(1).Col = 4 Then
            If NulosN(TxtCosto2.Text) = 0 Then
                MsgBox "Ingrese el costo por Hora", vbExclamation, xTitulo
                TxtCosto2.SetFocus
                Fg1(1).Editable = flexEDNone
                Exit Sub
            End If
            Fg1(1).Editable = flexEDKbdMouse
        Else
            Fg1(1).Editable = flexEDNone
        End If
    ElseIf Index = 2 Then
        If Fg1(2).Col = 5 Or Fg1(2).Col = 6 Or Fg1(2).Col = 12 Then
            Fg1(2).Editable = flexEDKbdMouse
        Else
            Fg1(2).Editable = flexEDNone
        End If
    End If
End Sub

Private Sub cb1_Click(Index As Integer)
    If QueHace <> 3 Then Exit Sub
    Dim xCampos() As String
    Dim nCampoBusca As String
    Dim nSQL As String
    Dim nTitulo As String
    On Error GoTo error
    Select Case Index
        Case 0 ' tarea
            ReDim xCampos(4, 4) As String
            xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":     xCampos(0, 2) = "4500":    xCampos(0, 3) = "C"
            xCampos(1, 0) = "Abrev":        xCampos(1, 1) = "nomcorto":   xCampos(1, 2) = "2300":    xCampos(1, 3) = "C"
            xCampos(2, 0) = "Diverso":      xCampos(2, 1) = "diverso":    xCampos(2, 2) = "700":     xCampos(2, 3) = "C"
            xCampos(3, 0) = "Id":           xCampos(3, 1) = "id":         xCampos(3, 2) = "600":     xCampos(3, 3) = "N"
            
            nSQL = "SELECT pro_tareas.id, pro_tareas.descripcion AS nombre, pro_tareas.id AS cod, pro_tareas.codigo, pro_tareas.abrev AS nomcorto, mae_unidades.id AS idunimed, mae_unidades.abrev, IIf([pro_tareas].[diverso]=-1,'Si','No') AS diverso " _
                    + vbCr + " FROM mae_unidades RIGHT JOIN pro_tareas ON mae_unidades.id = pro_tareas.idunimed  "

            nTitulo = "Buscando Tareas"
    End Select
    
    Dim RstTmp As New ADODB.Recordset
    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio

    If RstTmp.State = 0 Then GoTo SALIR
    If RstTmp.EOF = True Or RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR

    lbl_cod1(Index).Tag = lbl_cod1(Index).Caption

    txt_cb1(Index).Text = NulosC(RstTmp.Fields(0))         ' TEXTO A MOSTRAR
    lbl_cb1(Index).Caption = NulosC(RstTmp.Fields(1))      ' NOMBRE
    lbl_cod1(Index).Caption = NulosN(RstTmp.Fields(2))     ' CODIGO
    lbl_cb1(Index).ToolTipText = NulosC(RstTmp.Fields(1))  ' NOMBRE
    
    ' cargar datos de costo
    pDatosCostoCargar

SALIR:
    Set RstTmp = Nothing
    Exit Sub

error:
    Me.MousePointer = vbDefault
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "cb1_Click(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb1_Change(Index As Integer)
    If QueHace = 3 Then Exit Sub
    If txt_cb1(Index).Text = "" Then
        Me.lbl_cb1(Index).Caption = ""
        Me.lbl_cod1(Index).Caption = ""
    End If
    
End Sub

Private Sub txt_cb1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If txt_cb1(Index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb1_Click Index
        Exit Sub
    End If
End Sub

Private Sub txt_cb1_Validate(Index As Integer, Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    If txt_cb(Index).Text = "" Then Exit Sub
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    
    Select Case Index
        Case 0 ' area
            nSQL = "SELECT pro_tareas.id, pro_tareas.descripcion AS nombre, pro_tareas.id AS cod, pro_tareas.codigo, pro_tareas.abrev AS nomcorto, mae_unidades.id AS idunimed, mae_unidades.abrev, IIf([pro_tareas].[diverso]=-1,'Si','No') AS diverso " _
                + vbCr + " FROM mae_unidades RIGHT JOIN pro_tareas ON mae_unidades.id = pro_tareas.idunimed  " _
                + vbCr + " WHERE pro_tareas.id = " & NulosN(txt_cb(Index).Text) & " ;"

    End Select

    If xCon.State = 0 Then GoTo SALIR
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo SALIR

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cb1(Index).Text = NulosC(RstTmp.Fields(0))        ' TEXTO A MOSTRAR
        lbl_cb1(Index).Caption = NulosC(RstTmp.Fields(1))     ' NOMBRE
        lbl_cod1(Index).Caption = NulosN(RstTmp.Fields(2))    ' CODIGO
        lbl_cb1(Index).ToolTipText = NulosC(RstTmp.Fields(1)) ' NOMBRE
    Else
        txt_cb1(Index).Text = "":    lbl_cb1(Index).Caption = "":    lbl_cod1(Index).Caption = ""
    End If
    
    pDatosCostoCargar
    Set RstTmp = Nothing
    Exit Sub

error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "txt_cb1_Validate (" + CStr(Index) + ")"
    Exit Sub

SALIR:
    Set RstTmp = Nothing
    txt_cb(Index).Text = ""
End Sub

Private Sub pDatosCostoCargar()
        Dim Rst As New ADODB.Recordset
        Dim nSQL As String
        Dim nSQLTarea As String
        If NulosN(txt_cb1(0).Text) <> 0 Then
            nSQLTarea = " and pro_costodet.idtar = " & NulosN(txt_cb1(0).Text)
        End If
        
        ' cargando datos de las tareas directas
        nSQL = "SELECT pro_costodet.idcos, pro_costodet.corr, pro_costo.idref AS idrec, pro_costodet.idtar, pro_costodet.idunimed, alm_inventario.descripcion AS producto, pro_receta.codrec, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.cant, pro_costodet.costo, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio " _
            + vbCr + " FROM (alm_inventario RIGHT JOIN (pro_receta RIGHT JOIN pro_costo ON pro_receta.id = pro_costo.idref) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (mae_unidades INNER JOIN (pro_tareas INNER JOIN pro_costodet ON pro_tareas.id = pro_costodet.idtar) ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos " _
            + vbCr + " Where (((pro_costodet.idunimed) <> 7) And ((pro_costo.Tipo) = 1) And ((pro_costo.activo) = -1)) " & nSQLTarea _
            + vbCr + " ORDER BY alm_inventario.descripcion, pro_receta.codrec, pro_costo.tipo; "
    
        RST_Busq Rst, nSQL, xCon
        Agregando = True

        Fg1(0).Rows = 1
        If Rst.RecordCount <> 0 Then
            Do While Not Rst.EOF
                Fg1(0).Rows = Fg1(0).Rows + 1
                Fg1(0).TextMatrix(Fg1(0).Rows - 1, 1) = NulosC(Rst("producto"))
                Fg1(0).TextMatrix(Fg1(0).Rows - 1, 2) = NulosC(Rst("codrec"))
                Fg1(0).TextMatrix(Fg1(0).Rows - 1, 3) = NulosC(Rst("tarea"))
                Fg1(0).TextMatrix(Fg1(0).Rows - 1, 4) = NulosC(Rst("abrev"))
                Fg1(0).TextMatrix(Fg1(0).Rows - 1, 5) = NulosN(Rst("cant"))
                Fg1(0).TextMatrix(Fg1(0).Rows - 1, 6) = NulosN(Rst("costo"))
                Fg1(0).TextMatrix(Fg1(0).Rows - 1, 7) = NulosN(Rst("idrec"))
                Fg1(0).TextMatrix(Fg1(0).Rows - 1, 8) = NulosN(Rst("idtar"))
                Fg1(0).TextMatrix(Fg1(0).Rows - 1, 9) = NulosN(Rst("idunimed"))
                Fg1(0).TextMatrix(Fg1(0).Rows - 1, 10) = NulosN(Rst("idcos"))
                Fg1(0).TextMatrix(Fg1(0).Rows - 1, 11) = NulosN(Rst("corr"))
                If NulosN(Rst("costo")) = 0 Then
                    GRID_COLOR_FONDO Fg1(0), Fg1(0).Rows - 1, 6, Fg1(0).Rows - 1, 6, vbRed
                End If
                Rst.MoveNext
            Loop
        End If
        Set Rst = Nothing
        
        nSQL = "SELECT pro_costodet.idcos,pro_costodet.corr, pro_costodet.idtar, pro_costodet.idunimed, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.cant, pro_costodet.costo, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio " _
            + vbCr + " FROM pro_tareas INNER JOIN (pro_costo INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costo.idref " _
            + vbCr + " Where (((pro_costodet.idunimed) <> 7) And ((pro_costo.Tipo) = 2) And ((pro_costo.activo) = -1)) " & nSQLTarea _
            + vbCr + " ORDER BY pro_tareas.descripcion, pro_costo.tipo; "

        RST_Busq Rst, nSQL, xCon
        
        Fg1(1).Rows = 1
        If Rst.RecordCount <> 0 Then
            Do While Not Rst.EOF
                Fg1(1).Rows = Fg1(1).Rows + 1
                Fg1(1).TextMatrix(Fg1(1).Rows - 1, 1) = NulosC(Rst("tarea"))
                Fg1(1).TextMatrix(Fg1(1).Rows - 1, 2) = NulosC(Rst("abrev"))
                Fg1(1).TextMatrix(Fg1(1).Rows - 1, 3) = NulosN(Rst("cant"))
                Fg1(1).TextMatrix(Fg1(1).Rows - 1, 4) = NulosN(Rst("costo"))
                Fg1(1).TextMatrix(Fg1(1).Rows - 1, 5) = NulosN(Rst("idtar"))
                Fg1(1).TextMatrix(Fg1(1).Rows - 1, 6) = NulosN(Rst("idunimed"))
                Fg1(1).TextMatrix(Fg1(1).Rows - 1, 7) = NulosN(Rst("idcos"))
                Fg1(1).TextMatrix(Fg1(1).Rows - 1, 8) = NulosN(Rst("corr"))
                
                If NulosN(Rst("costo")) = 0 Then
                    GRID_COLOR_FONDO Fg1(1), Fg1(1).Rows - 1, 4, Fg1(1).Rows - 1, 4, vbRed
                End If
                Rst.MoveNext
            Loop
        End If
        Set Rst = Nothing
        Agregando = False
End Sub

Private Sub pDatosCostoGrabar()
    Dim mRow&
    If TabOne2.CurrTab = 0 Then
        If Fg1(0).Rows = 1 Then
            MsgBox "No hay Lista de tareas con Productos", vbExclamation, xTitulo
            Exit Sub
        End If
    Else
        If Fg1(1).Rows = 1 Then
            MsgBox "No hay Lista de tareas Diversos", vbExclamation, xTitulo
            Exit Sub
        End If
    End If
    
    If MsgBox("Seguro desea continuar", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    On Error GoTo error
    xCon.BeginTrans
    If TabOne2.CurrTab = 0 Then
        For mRow = 1 To Fg1(0).Rows - 1
            xCon.Execute "update pro_costodet set cant= " & NulosN(Fg1(0).TextMatrix(mRow, 5)) & ", costo = " & NulosN(Fg1(0).TextMatrix(mRow, 6)) & " where idcos= " & NulosN(Fg1(0).TextMatrix(mRow, 10)) & "  and corr = " & NulosN(Fg1(0).TextMatrix(mRow, 11)) & " and idtar = " & NulosN(Fg1(0).TextMatrix(mRow, 8)) & " and idunimed = " & NulosN(Fg1(0).TextMatrix(mRow, 9))
        Next mRow
    Else
        For mRow = 1 To Fg1(1).Rows - 1
            xCon.Execute "update pro_costodet set cant= " & NulosN(Fg1(1).TextMatrix(mRow, 3)) & ", costo = " & NulosN(Fg1(1).TextMatrix(mRow, 4)) & " where idcos= " & NulosN(Fg1(1).TextMatrix(mRow, 7)) & "  and corr = " & NulosN(Fg1(1).TextMatrix(mRow, 8)) & " and idtar = " & NulosN(Fg1(1).TextMatrix(mRow, 5)) & " and idunimed = " & NulosN(Fg1(1).TextMatrix(mRow, 6))
        Next mRow
    End If
    
    xCon.CommitTrans
    MsgBox "Se grabaron las tareas " & IIf(TabOne2.CurrTab = 0, " Directas", " Diversas ") & " con éxito", vbInformation, xTitulo
    
    Exit Sub

error:
    xCon.RollbackTrans
    SHOW_ERROR Me.Name, "pDatosCostoGrabar"
End Sub

Private Sub pDatosCostoCalcular(Tipo As Integer)
    ' tipo =1 actualizar costo
    ' tipo =2 equivalencia de costo a hora
    
    Dim mRow&
    If QueHace <> 3 Then Exit Sub
    If Tipo = 1 Then
        If NulosN(TxtCosto2.Text) = 0 Then
            MsgBox "Ingrese el costo por Hora", vbExclamation, xTitulo
            TxtCosto2.SetFocus
            Exit Sub
        End If
    Else
        If NulosN(TxtCosto3.Text) = 0 Then
            MsgBox "Ingrese el costo por Hora", vbExclamation, xTitulo
            TxtCosto3.SetFocus
            Exit Sub
        End If
    End If
    
    ' actualizar el costo/unidad de acuerdo al pago por hora
    Agregando = True
    If Tipo = 1 Then
        If TabOne2.CurrTab = 0 Then              ' costo directo
            For mRow = 0 To Fg1(0).Rows - 1
                If NulosN(Fg1(0).TextMatrix(mRow, 5)) <> 0 Then
                    Fg1(0).TextMatrix(mRow, 6) = NulosN(TxtCosto2.Text) / NulosN(Fg1(0).TextMatrix(mRow, 5))
                End If
            Next mRow
        Else
            For mRow = 0 To Fg1(1).Rows - 1
                If NulosN(Fg1(1).TextMatrix(mRow, 3)) <> 0 Then
                    Fg1(1).TextMatrix(mRow, 4) = NulosN(TxtCosto2.Text) / NulosN(Fg1(1).TextMatrix(mRow, 3))
                End If
            Next mRow
        End If
    Else
        For mRow = 0 To Fg1(2).Rows - 1
            If NulosN(Fg1(2).TextMatrix(mRow, 5)) <> 0 Then
                Fg1(2).TextMatrix(mRow, 6) = NulosN(TxtCosto3.Text) / NulosN(Fg1(2).TextMatrix(mRow, 5))
            End If
        Next mRow
    End If
    Agregando = False

End Sub

Private Sub pic_Click(Index As Integer)
    If Index = 0 Then
        pHabilitarBotonEditor False, 1
    Else
        pHabilitarBotonEditor False, 2
    End If
End Sub

Private Sub pConvertAHoraCargar()
    Dim Rst  As New ADODB.Recordset
    Dim nSQL As String

    If TxtFecha(0).valor = "" Or TxtFecha(1).valor = "" Then
        MsgBox "Ingrese una fecha", vbExclamation, xTitulo
        If TxtFecha(0).valor = "" Then TxtFecha(0).SetFocus Else TxtFecha(1).SetFocus
        Exit Sub
    End If
    If CDate(TxtFecha(0).valor) > CDate(TxtFecha(1).valor) Then
        MsgBox "La fecha inicial es superior al Final", vbExclamation, xTitulo
        TxtFecha(0).SetFocus
        Exit Sub
    End If

    nSQL = "SELECT vwTarea.codigopk, vwTarea.tarea, vwTarea.producto, vwTarea.abrev, vwTarea.idtar, vwTarea.idrec, vwTarea.idunimed, vwcosto.canteo, vwcosto.costo,vwcosto.idcos,vwcosto.corr,vwcosto.paghor " _
        + vbCr + " FROM ( "
    nSQL = nSQL _
        + vbCr + " SELECT  IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk,pro_controltar.idarea, pro_controltardet.idtar, pro_controltardet.idrec, pro_controltardet.idunimed, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, mae_unidades.abrev " _
        + vbCr + " FROM pro_controltar INNER JOIN (alm_inventario RIGHT JOIN (((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr " _
        + vbCr + " WHERE (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltar.tipo)=2) AND ((pro_controltardet.tipo)=1)) "
    nSQL = nSQL _
        + vbCr + " UNION "
    nSQL = nSQL _
        + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk,pro_controltar.idarea, pro_controltardet.idtar, pro_controltardet.idrec, pro_controltardet.idunimed, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, mae_unidades.abrev " _
        + vbCr + " FROM pro_controltar INNER JOIN ((alm_inventario RIGHT JOIN (((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN pro_controltardetgr ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) ON pro_controltar.id = pro_controltardet.idctr " _
        + vbCr + " WHERE (((pro_controltardet.idtar)<>0) AND ((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltar.tipo)=2) AND ((pro_controltardet.tipo)=2) AND ((pro_controltardetgr.activo)=-1)) "
    nSQL = nSQL _
        + vbCr + " ) AS vwtarea "
    nSQL = nSQL _
        + vbCr + " Left Join " _
        + vbCr + " (SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.idcos,pro_costodet.corr,pro_costodet.paghor,pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden  " _
        + vbCr + "  FROM pro_tareas INNER JOIN ((alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
        + vbCr + " ) AS vwcosto"
    nSQL = nSQL _
        + vbCr + " ON vwtarea.codigopk = vwcosto.codigopk "
            
    nSQL = nSQL _
     + vbCr + " WHERE (((vwTarea.tarea) Is Not Null) AND ((vwTarea.idunimed)<>7)) " _
     + vbCr + " ORDER BY vwTarea.producto, vwTarea.tarea; "
    
    RST_Busq Rst, nSQL, xCon

    Agregando = True

    Fg1(2).Rows = 1
    If Rst.RecordCount <> 0 Then
        Do While Not Rst.EOF
            Fg1(2).Rows = Fg1(2).Rows + 1
            Fg1(2).TextMatrix(Fg1(2).Rows - 1, 1) = NulosC(Rst("tarea"))
            Fg1(2).TextMatrix(Fg1(2).Rows - 1, 2) = NulosC(Rst("producto"))
'            Fg1(2).TextMatrix(Fg1(2).Rows - 1, 3) = NulosC(rst("codrec"))
            Fg1(2).TextMatrix(Fg1(2).Rows - 1, 4) = NulosC(Rst("abrev"))
            Fg1(2).TextMatrix(Fg1(2).Rows - 1, 5) = NulosN(Rst("canteo"))
            Fg1(2).TextMatrix(Fg1(2).Rows - 1, 6) = NulosN(Rst("costo"))
            Fg1(2).TextMatrix(Fg1(2).Rows - 1, 7) = NulosN(Rst("idrec"))
            Fg1(2).TextMatrix(Fg1(2).Rows - 1, 8) = NulosN(Rst("idtar"))
            Fg1(2).TextMatrix(Fg1(2).Rows - 1, 9) = NulosN(Rst("idunimed"))
            Fg1(2).TextMatrix(Fg1(2).Rows - 1, 10) = NulosN(Rst("idcos"))
            Fg1(2).TextMatrix(Fg1(2).Rows - 1, 11) = NulosN(Rst("corr"))
            Fg1(2).TextMatrix(Fg1(2).Rows - 1, 12) = NulosN(Rst("paghor"))
            
            If NulosN(Rst("costo")) = 0 Then
                GRID_COLOR_FONDO Fg1(2), Fg1(2).Rows - 1, 6, Fg1(2).Rows - 1, 6, vbRed
            End If
            
            Rst.MoveNext
        Loop
    End If
    Set Rst = Nothing
    
    Agregando = False
End Sub

Private Sub pConvertAHoraGrabar()
    Dim mRow&
    
    If Fg1(2).Rows = 1 Then
        MsgBox "No hay Lista de tareas con Productos", vbExclamation, xTitulo
        TxtFecha(0).SetFocus
        Exit Sub
    End If
    
    If MsgBox("Seguro desea continuar", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    On Error GoTo error
    xCon.BeginTrans
    
    For mRow = 1 To Fg1(2).Rows - 1
        xCon.Execute "update pro_costodet set cant= " & NulosN(Fg1(2).TextMatrix(mRow, 5)) & ", costo = " & NulosN(Fg1(2).TextMatrix(mRow, 6)) & ", paghor= " & NulosN(Fg1(2).TextMatrix(mRow, 12)) & " where idcos= " & NulosN(Fg1(2).TextMatrix(mRow, 10)) & "  and corr = " & NulosN(Fg1(2).TextMatrix(mRow, 11)) & " and idtar = " & NulosN(Fg1(2).TextMatrix(mRow, 8)) & " and idunimed = " & NulosN(Fg1(2).TextMatrix(mRow, 9))
    Next mRow
    
    xCon.CommitTrans
    MsgBox "Se grabaron las tareas con éxito", vbInformation, xTitulo
    Exit Sub

error:
    xCon.RollbackTrans
    SHOW_ERROR Me.Name, "pDatosCostoGrabar"
End Sub
