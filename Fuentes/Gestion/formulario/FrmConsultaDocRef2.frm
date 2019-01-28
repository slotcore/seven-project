VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmConsultaDocRef2 
   Caption         =   "Gestion - Analisis x Documento de Referencia"
   ClientHeight    =   7875
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   12555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   12555
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   1635
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Width           =   11790
      _cx             =   20796
      _cy             =   2884
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
      Appearance      =   2
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   0
      FrontTabForeColor=   -2147483630
      Caption         =   "Inicio|Mas"
      Align           =   0
      CurrTab         =   1
      FirstTab        =   0
      Style           =   3
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
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   1545
         Left            =   345
         TabIndex        =   8
         Top             =   45
         Width           =   11400
         Begin VB.Frame Frame13 
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
            Height          =   600
            Left            =   30
            TabIndex        =   22
            Top             =   0
            Width           =   3765
            Begin VB.CommandButton CmdBusMon 
               Height          =   230
               Left            =   495
               Picture         =   "FrmConsultaDocRef2.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   270
               Width           =   210
            End
            Begin VB.TextBox TxtIdMon 
               Height          =   300
               Left            =   180
               MaxLength       =   1
               TabIndex        =   24
               Text            =   "TxtIdMon"
               Top             =   240
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
               Height          =   300
               Left            =   735
               TabIndex        =   25
               Top             =   240
               Width           =   2925
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "[ Doc. Ref ]"
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
            Height          =   1545
            Left            =   8490
            TabIndex        =   38
            Top             =   0
            Width           =   2880
            Begin VSFlex7Ctl.VSFlexGrid Fg5 
               Height          =   1230
               Left            =   60
               TabIndex        =   39
               Top             =   240
               Width           =   2640
               _cx             =   4657
               _cy             =   2170
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
               Rows            =   50
               Cols            =   3
               FixedRows       =   0
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmConsultaDocRef2.frx":0132
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
         Begin VB.Frame Frame9 
            Caption         =   "[ Ordenado Por ]"
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
            Height          =   930
            Left            =   30
            TabIndex        =   17
            Top             =   615
            Width           =   3765
            Begin VB.OptionButton OptSort2 
               Caption         =   "Nº de Documento"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   135
               TabIndex        =   21
               Top             =   470
               Width           =   1800
            End
            Begin VB.OptionButton OptSort1 
               Caption         =   "Fecha  de Emisión"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   120
               TabIndex        =   20
               Top             =   240
               Value           =   -1  'True
               Width           =   2130
            End
            Begin VB.OptionButton OptSort3 
               Caption         =   "Nº Registro"
               ForeColor       =   &H00000080&
               Height          =   195
               Left            =   135
               TabIndex        =   19
               Top             =   700
               Width           =   1650
            End
            Begin VB.OptionButton OptSort4 
               Caption         =   "Fch. Emisión y Nº de Documento"
               ForeColor       =   &H00000080&
               Height          =   615
               Left            =   2250
               TabIndex        =   18
               Top             =   180
               Width           =   1470
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "[ Cliente ]"
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
            Height          =   1545
            Left            =   3810
            TabIndex        =   32
            Top             =   0
            Width           =   4650
            Begin VSFlex7Ctl.VSFlexGrid Fg2 
               Height          =   1230
               Left            =   60
               TabIndex        =   33
               Top             =   240
               Width           =   4530
               _cx             =   7990
               _cy             =   2170
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
               Rows            =   50
               Cols            =   3
               FixedRows       =   0
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmConsultaDocRef2.frx":0182
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
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   1545
         Left            =   -12045
         TabIndex        =   9
         Top             =   45
         Width           =   11400
         Begin VB.Frame Frame3 
            Caption         =   "[ Origenes ]"
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
            Height          =   1515
            Left            =   4260
            TabIndex        =   36
            Top             =   0
            Width           =   3180
            Begin VSFlex7Ctl.VSFlexGrid Fg3 
               Height          =   1230
               Left            =   60
               TabIndex        =   37
               Top             =   240
               Width           =   2955
               _cx             =   5212
               _cy             =   2170
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
               Rows            =   6
               Cols            =   3
               FixedRows       =   0
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmConsultaDocRef2.frx":01D2
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
            Caption         =   "[ Tipo Doc. Referencia ]"
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
            Height          =   1515
            Left            =   7500
            TabIndex        =   34
            Top             =   0
            Width           =   2940
            Begin VSFlex7Ctl.VSFlexGrid Fg4 
               Height          =   1230
               Left            =   60
               TabIndex        =   35
               Top             =   240
               Width           =   2610
               _cx             =   4604
               _cy             =   2170
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
               Cols            =   3
               FixedRows       =   0
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmConsultaDocRef2.frx":0287
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
         Begin VB.Frame Frame6 
            Caption         =   "[ Tipo Reporte ]"
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
            Height          =   900
            Left            =   1680
            TabIndex        =   26
            Top             =   600
            Width           =   2445
            Begin VB.OptionButton Option1 
               Caption         =   "Resumido"
               Height          =   195
               Left            =   30
               TabIndex        =   31
               Top             =   285
               Width           =   1020
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Detallado"
               Height          =   195
               Left            =   30
               TabIndex        =   30
               Top             =   555
               Width           =   990
            End
            Begin VB.Frame Frame8 
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
               ForeColor       =   &H00C00000&
               Height          =   570
               Left            =   1065
               TabIndex        =   27
               Top             =   240
               Width           =   1320
               Begin VB.OptionButton Option3 
                  Caption         =   "Con I.G.V."
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
                  Left            =   15
                  TabIndex        =   29
                  Top             =   45
                  Width           =   1230
               End
               Begin VB.OptionButton Option4 
                  Caption         =   "Sin I.G.V."
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
                  Left            =   15
                  TabIndex        =   28
                  Top             =   315
                  Width           =   1200
               End
            End
         End
         Begin VB.Frame Frame15 
            Caption         =   "[ Seleccionar Fecha ]"
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
            Left            =   30
            TabIndex        =   14
            Top             =   0
            Width           =   4095
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
               Height          =   300
               Left            =   735
               TabIndex        =   0
               Top             =   210
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
               Valor           =   "11/09/2008"
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
               Height          =   300
               Left            =   2700
               TabIndex        =   1
               Top             =   210
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
               Valor           =   "11/09/2008"
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Hasta"
               Height          =   195
               Index           =   2
               Left            =   2145
               TabIndex        =   16
               Top             =   255
               Width           =   420
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Desde"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   15
               Top             =   270
               Width           =   465
            End
         End
         Begin VB.Frame Frame10 
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
            Height          =   900
            Left            =   30
            TabIndex        =   10
            Top             =   600
            Width           =   1590
            Begin VB.OptionButton OptFch 
               Caption         =   "Fch Doc"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   13
               Top             =   240
               Width           =   1140
            End
            Begin VB.OptionButton OptFch 
               Caption         =   "Fch Reg"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   12
               Top             =   460
               Width           =   1125
            End
            Begin VB.OptionButton OptFch 
               Caption         =   "Fch Doc Ref"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   11
               Top             =   680
               Value           =   -1  'True
               Width           =   1425
            End
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11040
      Top             =   -90
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
            Picture         =   "FrmConsultaDocRef2.frx":02EC
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDocRef2.frx":0830
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDocRef2.frx":0BC2
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDocRef2.frx":0D46
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDocRef2.frx":119A
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDocRef2.frx":12B2
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDocRef2.frx":17F6
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDocRef2.frx":1D3A
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDocRef2.frx":1E4E
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDocRef2.frx":1F62
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDocRef2.frx":23B6
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDocRef2.frx":2522
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDocRef2.frx":2A6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDocRef2.frx":2D84
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaDocRef2.frx":3116
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   885
      Left            =   2580
      TabIndex        =   3
      Top             =   3570
      Visible         =   0   'False
      Width           =   6180
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   90
         TabIndex        =   4
         Top             =   390
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   6165
         Y1              =   870
         Y2              =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   6150
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   6165
         X2              =   6165
         Y1              =   15
         Y2              =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   840
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando Cta. Cte."
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
         Height          =   180
         Left            =   180
         TabIndex        =   5
         Top             =   90
         Width           =   1845
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   270
         Left            =   60
         Top             =   60
         Width           =   6075
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   5190
      Left            =   30
      TabIndex        =   2
      Top             =   2010
      Width           =   11745
      _cx             =   20717
      _cy             =   9155
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmConsultaDocRef2.frx":34A8
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
      TabIndex        =   6
      Top             =   0
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "1ra Forma"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "2da Forma"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Configurar Formatos"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu_01 
         Caption         =   "Agregar Cliente"
      End
      Begin VB.Menu menu_02 
         Caption         =   "-"
      End
      Begin VB.Menu menu_03 
         Caption         =   "Eliminar Cliente"
      End
   End
   Begin VB.Menu menu1 
      Caption         =   "menu1"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Agregar Documento"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar Documento"
      End
   End
End
Attribute VB_Name = "FrmConsultaDocRef2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SeEjecuto As Boolean

Dim xTotCom As Double 'Acumular importes x Orden - Compras
Dim xTotVen As Double 'Acumular importes x Orden - Ventas
Dim xTotLGD As Double 'Acumular importes x Orden - Lgd
Dim xTotRee As Double 'Acumular importes x Orden - Reembolsables
Dim xTotLet As Double 'Acumular importes x Orden - Letras/Abonos

Dim xTotComTot As Double 'Acumular importes x Cliente - Compras
Dim xTotVenTot As Double 'Acumular importes x Cliente - Ventas
Dim xTotLGDTot As Double 'Acumular importes x Cliente - Lgd
Dim xTotReeTot As Double 'Acumular importes x Cliente - Reembolsables
Dim xTotLetTot As Double 'Acumular importes x Cliente - Letras/Abonos


Private Sub pExportar(Optional band As Boolean = False)
    Dim xFun As New SGI2_funciones.Formularios
    Dim nTitulo As String
    Dim nPeriodo As String


    If Option1.Value = True Then '--resumen
        nTitulo = "Analisis x Documento de Referencia - Resumen"
    Else
        nTitulo = "Analisis x Documento de Referencia - Detallado"
    End If
    
    nPeriodo = "DEL " + TxtFchIni.Valor + " AL " + TxtFchFin.Valor
    
    If band = True Then '--formato rapido

        GRID_EXPORTAR_MSEXCELTMP Fg1, xCon, flexFileCustomText, True, nTitulo, nPeriodo, "Expresado en : " & LblMoneda.Caption

    Else
        xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1, nTitulo, nPeriodo, "", "Analisis x Documento de Referencia"
    End If
    

    Set xFun = Nothing
End Sub

Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    SelCliente
End Sub

Private Sub SelCliente()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(4, 4) As String
    Dim nSQLFiltro As String
    
    xCampos(0, 0) = "Cliente":       xCampos(0, 1) = "nombre":           xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "T.D.":          xCampos(1, 1) = "abrev":            xCampos(1, 2) = "800":          xCampos(1, 3) = "C"
    xCampos(2, 0) = "Nº Documento":  xCampos(2, 1) = "numruc":           xCampos(2, 2) = "1500":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Tipo Empresa":  xCampos(3, 1) = "tipemp":           xCampos(3, 2) = "1500":         xCampos(3, 3) = "C"
    
    
    '--Generar filtro de clientes; si seleccionaron en las lista
    '--Solamente se mostraran las ordenes de los clientes seleccionados; si no hay clientes seleccionado se muestra todo
    nSQLFiltro = GRID_GENERAR_SQL_ID(Fg2, 2, " AND mae_cliente.id", "NOT IN", True)
    
    
    xform.SQLCad = "SELECT mae_cliente.nombre, mae_dociden.abrev, mae_tipoempresa.descripcion AS tipemp, mae_cliente.numruc, " _
        & " mae_cliente.id FROM (mae_cliente LEFT JOIN mae_dociden ON mae_cliente.idtipdoc = mae_dociden.id) LEFT JOIN mae_tipoempresa " _
        & " ON mae_cliente.tipper = mae_tipoempresa.id Where (((mae_cliente.activo) = -1)) " & nSQLFiltro & "ORDER BY mae_cliente.nombre"
    
    xform.Titulo = "Buscando Cliente"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(xRs("nombre"))
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = xRs("id")
            
            If Fg2.TextMatrix(Fg2.Rows - 1, 1) <> "" Then
                Fg2.Rows = Fg2.Rows + 1
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Fg2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        menu_01_Click
    End If
    
    If KeyCode = 46 Then
        menu_03_Click
    End If
End Sub

Private Sub Fg2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu Menu
    End If
End Sub

Private Sub Fg3_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Then
        
    End If
End Sub

Private Sub Fg3_EnterCell()
    If Fg3.Col = 2 Then
        Fg3.Editable = flexEDKbdMouse
    Else
        Fg3.Editable = flexEDNone
    End If
End Sub

Private Sub Fg4_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Then
        If NulosN(Fg4.TextMatrix(Row, Col)) = -1 Then
'            Fg3.Rows = 0
'            Fg3.Rows = 7
'            Fg3.TextMatrix(0, 1) = "VENTAS"
'            Fg3.TextMatrix(1, 1) = "COMPRAS"
'            Fg3.TextMatrix(2, 1) = "HONORARIOS"
'            Fg3.TextMatrix(3, 1) = "REEMBOLSABLES"
'            Fg3.TextMatrix(4, 1) = "L.G.D. / L.G.C."
'            Fg3.TextMatrix(5, 1) = "LETRAS/ABONOS"

'
'            Fg3.TextMatrix(0, 2) = -1
'            Fg3.TextMatrix(1, 2) = -1
'            Fg3.TextMatrix(2, 2) = -1
'            Fg3.TextMatrix(3, 2) = -1
'            Fg3.TextMatrix(4, 2) = -1
'            Fg3.TextMatrix(5, 2) = -1

            OptFch(2).Enabled = True
        Else
'            Fg3.Rows = 1
            OptFch(0).Value = True
            OptFch(2).Enabled = False
        End If
'        Option2.Value = True
'        Option2_Click
    End If
End Sub

Private Sub Fg4_EnterCell()
    If Fg4.Col = 2 Then
        Fg4.Editable = flexEDKbdMouse
    Else
        Fg4.Editable = flexEDNone
    End If
End Sub


Private Sub Fg5_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        menu1_1_Click
    End If
    
    If KeyCode = 46 Then
        menu1_3_Click
    End If
End Sub

Private Sub Form_Load()
'    Me.WindowState = 2
    SeEjecuto = False
    Frame6.BackColor = &H8000000F
    
    SetearCuadricula Fg1, 7, xCon, 2, 2, False
    
    Fg1.BackColor = &HE2FEFB
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.BackColorSel = &H80&
    Fg1.Editable = flexEDNone
        
    Fg2.BackColor = &HE2FEFB
    Fg2.Rows = 0
    Fg2.Rows = Fg2.Rows + 1
    Fg2.ColComboList(1) = "|..."
    Fg2.SelectionMode = flexSelectionByRow
    Fg2.ColWidth(2) = 0
    Fg2.Editable = flexEDKbdMouse
    
    Fg5.BackColor = &HE2FEFB
    Fg5.Rows = 0
    Fg5.Rows = Fg5.Rows + 1
    Fg5.ColComboList(1) = "|..."
    Fg5.SelectionMode = flexSelectionByRow
    Fg5.ColWidth(2) = 0
    Fg5.Editable = flexEDKbdMouse
    
    Dim A As Integer
    For A = 0 To 4
        Fg3.TextMatrix(A, 2) = -1
    Next A
        
    Fg3.Editable = flexEDNone
    Fg3.SelectionMode = flexSelectionByRow
    
    Fg4.TextMatrix(0, 2) = -1
    Fg4.Editable = flexEDNone
    Fg4.SelectionMode = flexSelectionByRow
    
    
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
End Sub
Private Sub Form_Activate()
    If SeEjecuto = False Then
        
        TxtFchIni.Valor = Date
        TxtFchFin.Valor = Date

        TxtIdMon.Text = 1
        TxtIdMon_Validate False
        
        TabOne1.CurrTab = 0
        
        Option2.Value = True   ' ponemos la opcion detallado por defecto
        Option3.Value = True   ' ponemos la opcion con IGV por defecto
        
        OptSort1.Value = True '--orden por fecha de emision x defecto
    
        TxtFchIni.Valor = ""
        TxtFchFin.Valor = ""
        
        TxtFchIni.SetFocus
        
        SeEjecuto = True
    End If
End Sub

Sub TotalizarCliente()
    '===================================================================================================
    'creado: xx/02/10 Por Enrique Pollongo
    '
    'Propósito: mostrar el acumulado por cliente
    '
    'Entradas:  Indice = Ninguno
    '
    'Resultados:    Acumulados en pantalla segun columnas
    '
    'Modificado:    22/09/10 Johan Castro
    '               Considerar la exprexión a una moneda, segun selecciona del usuario
    '===================================================================================================
    
    
    'totalizamos
    Fg1.Rows = Fg1.Rows + 1
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 16, &H800000, True, &HE2FEFB, "TOTAL CLIENTE ==> "
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 17, &H80000012, True, &HE2FEFB, Format(xTotComTot, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 18, &H80000012, True, &HE2FEFB, Format(xTotVenTot, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 19, &H80000012, True, &HE2FEFB, Format(xTotReeTot, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 20, &H80000012, True, &HE2FEFB, Format(xTotLGDTot, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 21, &H80000012, True, &HE2FEFB, Format(xTotLetTot, FORMAT_MONTO)
    
    Fg1.Rows = Fg1.Rows + 1
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 16, &H800000, True, &HE2FEFB, "SALDO CLIENTE ==> "
    
    If (xTotVenTot - xTotComTot) < 0 Then
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 18, &HC0&, True, &HE2FEFB, Format(xTotVenTot - xTotComTot, FORMAT_MONTO)
    Else
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 18, &HC00000, True, &HE2FEFB, Format(xTotVenTot - xTotComTot, FORMAT_MONTO)
    End If
    
    '----------------------------------------------------------------------------------------------------------------
    If (xTotLGDTot - xTotReeTot) < 0 Then
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 20, &HC0&, True, &HE2FEFB, Format(xTotLGDTot - xTotReeTot, FORMAT_MONTO)
    Else
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 20, &HC00000, True, &HE2FEFB, Format(xTotLGDTot - xTotReeTot, FORMAT_MONTO)
    End If
    
    'let-(lgd-vta)
    If (xTotLetTot - (xTotVenTot + xTotLGDTot)) < 0 Then
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 21, &HFF&, True, &HE2FEFB, Format(xTotLetTot - (xTotVenTot + xTotLGDTot), FORMAT_MONTO)
    Else
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 21, &HFF0000, True, &HE2FEFB, Format(xTotLetTot - (xTotVenTot + xTotLGDTot), FORMAT_MONTO)
    End If
    
    '--reinciar varialbles
    xTotComTot = 0
    xTotVenTot = 0
    xTotLGDTot = 0
    xTotReeTot = 0
    xTotLetTot = 0
    
End Sub

Sub Totalizar()
    '===================================================================================================
    'creado: xx/02/10 Por Enrique Pollongo
    '
    'Propósito: mostrar el acumulado por documento de referencia
    '
    'Entradas:  Indice = Ninguno
    '
    'Resultados:    Acumulados en pantalla segun columnas
    '
    'Modificado:    22/09/10 Johan Castro
    '               Considerar la exprexión a una moneda, segun selecciona del usuario
    '               17/12/10 Johan Castro
    '               La presentacion de los acumulados total y saldo por orden no coinciden con la columna,
    '               se corrige en su orden correcto.
    '===================================================================================================
    
    'totalizamos
    Fg1.Rows = Fg1.Rows + 1
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 16, &H800000, True, &HE2FEFB, "TOTAL ORDEN ==> "
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 17, &H80000012, True, &HE2FEFB, Format(xTotCom, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 18, &H80000012, True, &HE2FEFB, Format(xTotVen, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 19, &H80000012, True, &HE2FEFB, Format(xTotRee, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 20, &H80000012, True, &HE2FEFB, Format(xTotLGD, FORMAT_MONTO)
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 21, &H80000012, True, &HE2FEFB, Format(xTotLet, FORMAT_MONTO)
    
    Fg1.Rows = Fg1.Rows + 1
    
    FORMATO_CELDA Fg1, Fg1.Rows - 1, 16, &H800000, True, &HE2FEFB, "SALDO ORDEN ==> "
    
    If (xTotVen - xTotCom) < 0 Then
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 18, &HFF&, True, &HE2FEFB, Format(xTotVen - xTotCom, FORMAT_MONTO)
    Else
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 18, &HFF0000, True, &HE2FEFB, Format(xTotVen - xTotCom, FORMAT_MONTO)
    End If
    
    '--------------------------------------------------------------------------------------------------------------
    If (xTotLGD - xTotRee) < 0 Then
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 20, &HFF&, True, &HE2FEFB, Format(xTotLGD - xTotRee, FORMAT_MONTO)
    Else
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 20, &HFF0000, True, &HE2FEFB, Format(xTotLGD - xTotRee, FORMAT_MONTO)
    End If
    
    If (xTotLet - (xTotVen + xTotLGD)) < 0 Then
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 21, &HFF&, True, &HE2FEFB, Format(xTotLet - (xTotVen + xTotLGD), FORMAT_MONTO)
    Else
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 21, &HFF0000, True, &HE2FEFB, Format(xTotLet - (xTotVen + xTotLGD), FORMAT_MONTO)
    End If
    
    '--reiniciar variables
    xTotCom = 0
    xTotVen = 0
    xTotLGD = 0
    xTotRee = 0
    xTotLet = 0
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub

    If Me.Height > 3000 Then
        Fg1.Top = 2010
        Fg1.Width = Me.Width - 150
        Fg1.Height = Me.Height - 2400
    End If

End Sub

Private Sub menu_01_Click()
    If Fg2.Rows = 0 Then
        Fg2.Rows = Fg2.Rows + 1
        Exit Sub
    End If
    
    If NulosC(Fg2.TextMatrix(Fg2.Rows - 1, 1)) <> "" Then
        Fg2.Rows = Fg2.Rows + 1
    End If
    SelCliente
End Sub

Private Sub menu_03_Click()
    If Fg2.Row < 0 Then
        MsgBox "Seleccionar fila correcta", vbInformation, xTitulo
        Exit Sub
    End If
    If Fg2.Rows <> 0 Then
        Fg2.RemoveItem Fg2.Row
    End If
    If Fg2.Rows = 0 Then
        Fg2.Rows = Fg2.Rows + 1
    End If
End Sub

Private Sub menu1_1_Click()
    If Fg5.Rows = 0 Then
        Fg5.Rows = Fg5.Rows + 1
        Exit Sub
    End If
    
    If NulosC(Fg5.TextMatrix(Fg5.Rows - 1, 1)) <> "" Then
        Fg5.Rows = Fg5.Rows + 1
    End If
    
    SelDocReferencia
    
End Sub

Private Sub menu1_3_Click()
    If Fg5.Row < 0 Then
        MsgBox "Seleccionar fila correcta", vbInformation, xTitulo
        Exit Sub
    End If
    If Fg5.Rows <> 0 Then
        Fg5.RemoveItem Fg5.Row
    End If
    If Fg5.Rows = 0 Then
        Fg5.Rows = Fg5.Rows + 1
    End If
End Sub

Private Sub Option1_Click()
'    If NulosN(Fg4.TextMatrix(0, 2)) = -1 Then
        If Option1.Value = True Then
            SetearCuadricula Fg1, 7, xCon, 2, 1, False
            '--indicar la moneda del reporte
            UNIR_CELDAS Fg1, 0, 6, 0, 10, "EXPRESADO EN " & UCase(LblMoneda.Caption), , True
        End If
'    Else
'        If Option1.Value = True Then SetearCuadricula Fg1, 7, xCon, 2, 3, False
'    End If
End Sub

Private Sub Option2_Click()
'    If NulosN(Fg4.TextMatrix(0, 2)) = -1 Then
        If Option2.Value = True Then
            SetearCuadricula Fg1, 7, xCon, 2, 2, False
            '--indicar la moneda del reporte
            UNIR_CELDAS Fg1, 0, 17, 0, 21, "EXPRESADO EN " & UCase(LblMoneda.Caption), , True
        End If
'    Else
'        If Option2.Value = True Then SetearCuadricula Fg1, 7, xCon, 2, 4, False
'    End If
    
End Sub



Sub VerDetalle()
    '===================================================================================================
    'Creado: 22/09/10
    'Propósito: muestra consulta detallada en pantalla
    '
    'Entradas:  Indice = Ninguno
    '
    'Resultados: Consulta en pantalla segun parametros ingresados por usuario
    '===================================================================================================
   
    Dim Rst As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLSub As String '--Sentencia SQL para identificar una subconsulta; está a nivel de detalle
    Dim nSQLSort As String '--Sentencia SQL para aplicar orden a la consulta
    Dim A As Double
   
    Dim xCliente As String '--Util para generar los grupos por cliente
    Dim NumDocRef As String '--Util para generar los grupos por numero de referencia
   
    '--Generar la sub consulta
    nSQLSub = GenerarConsulta()
    
    '--aplicar el orden
    If OptSort1.Value = True Then
        nSQLSort = ",det.fchdoc "
    ElseIf OptSort2.Value = True Then
        nSQLSort = ",det.numdocumento "
    ElseIf OptSort3.Value = True Then
        nSQLSort = ",det.registro "
    Else
        nSQLSort = ",det.fchdoc,det.numdocumento "
    End If
    
    '--verificar la moneda a expresar la consulta
    If NulosN(TxtIdMon.Text) = 1 Then '--moneda nacional
    
        nSQL = "SELECT det.modulo, det.refnomcli,det.refabrev,det.refnumdoc,det.reffchdoc,det.registro,det.numruc,det.nombre,det.abrev,det.numdocumento,det.fchdoc,det.simbolo,det.tipcam,det.impreal,det.glosa, " _
                + vbCr + " det.compsol as impcomp ,det.vtasol as impvta,det.reemsol as impreem,det.lgdsol as implgd,det.letsol as implet " _
                + vbCr + " FROM ( " _
                + vbCr + nSQLSub _
                + vbCr + " ) AS det " _
                + vbCr + " ORDER BY det.refnomcli, det.refnumdoc, det.reffchdoc " & nSQLSort
                
    Else '--moneda extrnajera
        nSQL = "SELECT det.modulo,det.refnomcli,det.refabrev,det.refnumdoc,det.reffchdoc,det.registro,det.numruc,det.nombre,det.abrev,det.numdocumento,det.fchdoc,det.simbolo,det.tipcam,det.impreal,det.glosa, " _
                + vbCr + " det.compdol as impcomp ,det.vtadol as impvta,det.reemdol as impreem,det.lgddol as implgd,det.letdol as implet " _
                + vbCr + " FROM ( " _
                + vbCr + nSQLSub _
                + vbCr + " ) AS det " _
                + vbCr + " ORDER BY det.refnomcli, det.refnumdoc, det.reffchdoc " & nSQLSort
    End If
    
    '--cambiar cursor de espera del mouse
    Me.MousePointer = vbHourglass
    
    '--ejecutar la consulta
    RST_Busq Rst, nSQL, xCon
    
    If Rst.State = 0 Then GoTo LaCague:
    
    '--verificar si hay registros
    If Rst.RecordCount <> 0 Then
        
        '--centrar la barra de progreso
        Frame5.Left = (Me.Width - Frame5.Width) / 2
        Frame5.Top = (Me.Height - Frame5.Height) / 2
        '--obtener cantidad de registros
        ProgressBar1.Max = Rst.RecordCount
        '--mostrar la barra de progreso
        Frame5.Visible = True
        
        '--posicionar a la primera fila
        Rst.MoveFirst
        Fg1.Rows = 2
        NumDocRef = NulosC(Rst("refnumdoc"))
        '--reiniciar variables
        xTotCom = 0
        xTotVen = 0
        xTotLGD = 0
        xTotRee = 0
        xTotLet = 0

        
        xTotVenTot = 0
        xTotComTot = 0
        xTotLGDTot = 0
        xTotReeTot = 0
        xTotLetTot = 0
              
        For A = 1 To Rst.RecordCount
            ProgressBar1.Value = A
            
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(Rst("refnomcli"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(Rst("refabrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(Rst("refnumdoc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(NulosC(Rst("reffchdoc")), "dd/mm/yy")
            If IsNull(Rst("reffchdoc")) = False Then Fg1.TextMatrix(Fg1.Rows - 1, 5) = (CDate(TxtFchFin.Valor) - Rst("reffchdoc"))
            
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(Rst("modulo"))
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosC(Rst("registro"))
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosC(Rst("numruc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosC(Rst("nombre"))
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosC(Rst("abrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosC(Rst("numdocumento"))
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(Rst("fchdoc"), "dd/mm/yy")
            Fg1.TextMatrix(Fg1.Rows - 1, 13) = NulosC(Rst("simbolo"))
            Fg1.TextMatrix(Fg1.Rows - 1, 14) = Format(Rst("tipcam"), "0.000")
            Fg1.TextMatrix(Fg1.Rows - 1, 15) = Format(Rst("impreal"), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 16) = NulosC(Rst("glosa"))
            
            Fg1.TextMatrix(Fg1.Rows - 1, 17) = Format(NulosN(Rst("impcomp")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 18) = Format(NulosN(Rst("impvta")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(NulosN(Rst("impreem")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 20) = Format(NulosN(Rst("implgd")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 21) = Format(NulosN(Rst("implet")), FORMAT_MONTO)
            
            '--acumulando para mostrar resumen por orden
            xTotCom = xTotCom + NulosN(Rst("impcomp"))
            xTotVen = xTotVen + NulosN(Rst("impvta"))
            xTotRee = xTotRee + NulosN(Rst("impreem"))
            xTotLGD = xTotLGD + NulosN(Rst("implgd"))
            xTotLet = xTotLet + NulosN(Rst("implet"))
            
            '--acumulando los subtotales
            xTotComTot = xTotComTot + NulosN(Rst("impcomp"))
            xTotVenTot = xTotVenTot + NulosN(Rst("impvta"))
            xTotReeTot = xTotReeTot + NulosN(Rst("impreem"))
            xTotLGDTot = xTotLGDTot + NulosN(Rst("implgd"))
            xTotLetTot = xTotLetTot + NulosN(Rst("implet"))
                        
            xCliente = NulosC(Rst("refnomcli"))
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Totalizar
                Fg1.Rows = Fg1.Rows + 1
                TotalizarCliente
                Exit For
            End If
            
            ' si el numero de documento de referencia a cambiado
            If NumDocRef <> NulosC(Rst("refnumdoc")) Then
                Totalizar
                Fg1.Rows = Fg1.Rows + 1
            End If
            
            NumDocRef = NulosC(Rst("refnumdoc"))
            
            If xCliente <> Rst("refnomcli") Then
                TotalizarCliente
                Fg1.Rows = Fg1.Rows + 2
                xCliente = NulosC(Rst("refnomcli"))

            End If

        Next A
        
    End If
    
LaCague:
    Frame5.Visible = False
    Set Rst = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub Fg5_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    SelDocReferencia
End Sub

Private Sub SelDocReferencia()
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(4, 4) As String
    Dim nSQL As String
    Dim nSQLFiltro As String
    
    'Orden de Despacho
    xCampos(0, 0) = "Nº Documento":      xCampos(0, 1) = "numdoc":      xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Fch. Emi.":         xCampos(1, 1) = "fchemi":      xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Fch. Ven.":         xCampos(2, 1) = "fchven":      xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Cliente":           xCampos(3, 1) = "nombre":      xCampos(3, 2) = "4000":         xCampos(3, 3) = "C"
    
    '--Generar filtro de clientes; si seleccionaron en las lista
    '--Solamente se mostraran las ordenes de los clientes seleccionados; si no hay clientes seleccionado se muestra todo
    nSQLFiltro = GRID_GENERAR_SQL_ID(Fg2, 2, " WHERE mae_cliente.id", "IN", True)
    
    If nSQLFiltro = "" Then
        nSQLFiltro = GRID_GENERAR_SQL_ID(Fg5, 2, " WHERE var_ordendespacho.id", "NOT IN", True)
    Else
        nSQLFiltro = nSQLFiltro & GRID_GENERAR_SQL_ID(Fg5, 2, " AND var_ordendespacho.id", "NOT IN", True)
    End If
    
    
    nSQL = "SELECT 0 as xsel,var_ordendespacho.id, var_ordendespacho.numerodoc AS numdoc,mae_cliente.nombre, var_ordendespacho.idcli, var_ordendespacho.fchemi, var_ordendespacho.fchven  " _
            & " FROM var_ordendespacho LEFT JOIN mae_cliente ON var_ordendespacho.idcli = mae_cliente.id " _
            & nSQLFiltro
            
  
    CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), "Buscanco Documento de Referencia"
        
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            Do While Not xRs.EOF
                If Fg5.Rows = Fg5.FixedRows Then Fg5.Rows = Fg5.Rows + 1
                Fg5.Row = Fg5.Rows - 1
                
                If NulosN(Fg5.TextMatrix(Fg5.Rows - 1, 2)) <> 0 Then
                    Fg5.Rows = Fg5.Rows + 1
                    Fg5.Row = Fg5.Rows - 1
                End If
                
                Fg5.TextMatrix(Fg5.Rows - 1, 1) = NulosC(xRs("numdoc"))
                Fg5.TextMatrix(Fg5.Rows - 1, 2) = xRs("id")
                
                If NulosN(Fg5.TextMatrix(Fg5.Rows - 1, 2)) <> 0 Then Fg5.Rows = Fg5.Rows + 1
                
                xRs.MoveNext
                
            Loop
        End If
    End If
    
    Set xRs = Nothing

End Sub

Private Sub Fg5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu Menu1
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        '--limpiar datos
        Fg1.Rows = Fg1.FixedRows
        DoEvents
        
        '--posicionar en la primera pestaña
        TabOne1.CurrTab = 0
        DoEvents
        '--
        ' VERIFICAMOS QUE LOS DATOS NECESARIOS SEAN LOS CORRECTOS
        If NulosC(TxtFchIni.Valor) = "" Then
            MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchIni.SetFocus
            Exit Sub
        End If
        
        If NulosC(TxtFchFin.Valor) = "" Then
            MsgBox "No ha especificado la fecha de final", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtFchFin.SetFocus
            Exit Sub
        End If

        If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
            MsgBox "La fecha de inicio no puede ser mayor a la fecha final", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
            TxtFchIni.SetFocus
            Exit Sub
        End If
        
        '--VERIFICAMOS LA MONEDA
        If NulosN(TxtIdMon.Text) = 0 Then
            MsgBox "Falta especificar la moneda", vbInformation, xTitulo
            TabOne1.CurrTab = 1
            TxtIdMon.SetFocus
            Exit Sub
        End If
        '--
        
        'Procesar
        If Option1.Value = True Then '--ver resumen
            VerResumen
        Else '--ver detalle
            VerDetalle
        End If
        
    End If
    
    If Button.Index = 3 Then
        
        pExportar True
        
    End If
    
    If Button.Index = 4 Then
'        pImprimir
    End If
    
    If Button.Index = 5 Then
'        Configurar
    End If
    
    If Button.Index = 7 Then
        Unload Me
    End If
End Sub



Private Sub CmdBusMon_Click()
    
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre":      xCampos(0, 1) = "descripcion":     xCampos(0, 2) = "4500":      xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":   xCampos(1, 1) = "id":              xCampos(1, 2) = "500":      xCampos(1, 3) = "N"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, "SELECT * FROM mae_moneda ORDER BY descripcion ;", xCampos(), "Buscando Moneda", "descripcion", "descripcion", Principio
    
    If xRs.State = 0 Then GoTo salir
    If xRs.RecordCount = 0 Then GoTo salir
    TxtIdMon.Text = xRs("id") & ""
    LblMoneda.Caption = xRs("descripcion") & ""
    
     '--indicar la moneda del reporte
    If Option1.Value = True Then
        UNIR_CELDAS Fg1, 0, 6, 0, 10, "EXPRESADO EN " & UCase(LblMoneda.Caption), , True
    Else
        UNIR_CELDAS Fg1, 0, 17, 0, 21, "EXPRESADO EN " & UCase(LblMoneda.Caption), , True
    End If
    
salir:
    Set xRs = Nothing
End Sub

Private Sub TxtIdMon_Change()
    If Trim(TxtIdMon.Text) = "" Then TxtIdMon.Text = ""
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
        If NulosC(TxtIdMon.Text) = "" Then
            TxtIdMon.Text = ""
        End If
        
         '--indicar la moneda del reporte
        If Option1.Value = True Then
            UNIR_CELDAS Fg1, 0, 6, 0, 10, "EXPRESADO EN " & UCase(LblMoneda.Caption), , True
        Else
            UNIR_CELDAS Fg1, 0, 17, 0, 21, "EXPRESADO EN " & UCase(LblMoneda.Caption), , True
        End If
    End If
End Sub


Private Function GenerarConsulta() As String
    '===================================================================================================
    'creado: 24/09/10 Por Johan Castro
    'Propósito: Generar la consulta a nivel de detalle
    '
    'Entradas:  Ninguno
    '
    'Resultados: Consulta segun parametros indicados
    '
    'Modificado:03/11/10 Johan Castro
    '           Agregar consulta de abonos que hacen referencia a ordenes
    '           Considerar las devoluciones de los abonos
    '           05/11/10 Johan Castro
    '           Agregar los cargos x devolucion al cliente
    '
    '===================================================================================================

    Dim nSQL As String
    Dim nSQLFiltro As String
    Dim nSQLFiltroFch As String '--filtro de periodo
    Dim nSQLFilDocRef As String '--filtro para documento de referencia
    
    '--verificar si muestra documentos sin referencia
    If Fg4.TextMatrix(0, 2) = 0 Then
        '--sin referencia
        nSQLFiltro = " and var_ordendespacho.id is null "
    Else
        '--con referencia
        nSQLFiltro = " and var_ordendespacho.id is not null "
    End If
    
    '--filtro de periodo para
    nSQLFiltroFch = " and var_ordendespacho.fchemi between cdate('" & TxtFchIni.Valor & "') and cdate('" & TxtFchFin.Valor & "')"
    
    '--filtro para documentos de referencia
    nSQLFilDocRef = GRID_GENERAR_SQL_ID(Fg5, 2, " and var_ordendespacho.id", "IN", True)
    If nSQLFilDocRef = "" Then
        '--filtro para cliente
        nSQLFilDocRef = GRID_GENERAR_SQL_ID(Fg2, 2, " and var_ordendespacho.idcli", "IN", True)
    Else
        '--limpiar el filtro de fechas cuando se consulte documentos especificos
        nSQLFiltroFch = ""
    End If
    '--uniendo las condiciones
    nSQLFiltro = nSQLFiltro & nSQLFiltroFch & nSQLFilDocRef
    
    '--Ventas
    If NulosN(Fg3.TextMatrix(0, 2)) = -1 Then
           
        nSQL = "SELECT 'Ventas' as modulo, var_ordendespacho.id AS refid, mae_documento_1.abrev AS refabrev, mae_cliente_1.nombre AS refnomcli, var_ordendespacho.numerodoc AS refnumdoc, var_ordendespacho.fchemi AS reffchdoc, " _
            + vbCr + " vta_ventas.id AS iddoc, Left([vta_ventas].[numreg],2) & [mae_libros].[codsun] & Right([vta_ventas].[numreg],4) AS registro, mae_cliente.numruc, mae_cliente.nombre, mae_documento.abrev, [vta_ventas].[numser] & '-' & [vta_ventas].[numdoc] AS numdocumento, vta_ventas.fchdoc, vta_ventas.fchven, vta_ventas.glosa, mae_moneda.simbolo, " _
            + vbCr + " IIf([vta_ventas].[tc]=0,IIf([con_tc].[impven] Is Null,0,[con_tc].[impven]),[vta_ventas].[tc]) AS tipcam, IIf(vta_ventas.tipdoc=7,(-1)*[vta_ventas].[imptotdoc],[vta_ventas].[imptotdoc]) AS impreal, " _
            + vbCr + " IIf([vta_ventas].[idmon]=1,impreal,impreal*tipcam) AS vtasol, " _
            + vbCr + " 0 as compsol,0 as reemsol,0 as lgdsol,0 as letsol, " _
            + vbCr + " IIf([vta_ventas].[idmon]=2,impreal,IIf(tipcam=0,0,impreal/tipcam)) AS vtadol, " _
            + vbCr + " 0 as compdol,0 as reemdol,0 as lgddol,0 as letdol " _
            + vbCr + " FROM (((mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN (((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON mae_moneda.id = vta_ventas.idmon) ON mae_cliente.id = vta_ventas.idcli) " _
            + vbCr + " LEFT JOIN var_ordendespacho ON vta_ventas.iddocref2 = var_ordendespacho.id) LEFT JOIN mae_cliente AS mae_cliente_1 ON var_ordendespacho.idcli = mae_cliente_1.id) LEFT JOIN mae_documento AS mae_documento_1 ON var_ordendespacho.idtipdoc = mae_documento_1.id " _
            + vbCr + " WHERE ((vta_ventas.numreg)<>'000001') AND ((vta_ventas.anulado)=0) " & nSQLFiltro
            '+ vbCr + " WHERE (((vta_ventas.fchreg) Between CDate('01/01/10') And CDate('31/01/10')) AND (vta_ventas.numreg)<>'000001') AND ((vta_ventas.anulado)=0)) "
           
        '--fltro por fecha del documento
        If OptFch(0).Value = True Then nSQL = Replace(nSQL, "var_ordendespacho.fchemi between", "vta_ventas.fchdoc between")
        '--filtro por fecha de registro
        If OptFch(1).Value = True Then nSQL = Replace(nSQL, "var_ordendespacho.fchemi between", "vta_ventas.fchreg between")
        '--reporte sin igv
        If Option4.Value = True Then nSQL = Replace(nSQL, "[vta_ventas].[imptotdoc]", "(vta_ventas.impbru + vta_ventas.impbru2 + vta_ventas.impbru3 + vta_ventas.impinaf)")
        
    End If
    
    '--Compras
    If NulosN(Fg3.TextMatrix(1, 2)) = -1 Then
        '--verificar si hay union de consultas
        If nSQL <> "" Then nSQL = nSQL + vbCr + " UNION "
        
        nSQL = nSQL _
            + vbCr + " SELECT 'Compras' as modulo, var_ordendespacho.id AS refid, mae_documento_1.abrev AS refabrev, mae_cliente_1.nombre AS refnomcli, var_ordendespacho.numerodoc AS refnumdoc, var_ordendespacho.fchemi AS reffchdoc, " _
            + vbCr + " com_compras.id AS iddoc, Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Right([com_compras].[numreg],4) AS registro, mae_prov.numruc, mae_prov.nombre, mae_documento.abrev, [com_compras].[numser] & '-' & [com_compras].[numdoc] AS numdocumento, com_compras.fchdoc, com_compras.fchven, com_compras.glosa, mae_moneda.simbolo, " _
            + vbCr + " IIf([com_compras].[tc]=0,IIf([con_tc].[impven] Is Null,0,[con_tc].[impven]),[com_compras].[tc]) AS tipcam, IIf(com_compras.tipdoc=7,(-1)*[com_compras].[imptot],[com_compras].[imptot]) AS impreal, " _
            + vbCr + " 0 as vtasol, " _
            + vbCr + " IIf([com_compras].[idmon]=1,impreal,impreal*tipcam) AS compsol, " _
            + vbCr + " 0 as reemsol,0 as lgdsol,0 as letsol,0 as vtadol, " _
            + vbCr + " IIf([com_compras].[idmon]=2,impreal,IIf(tipcam=0,0,impreal/tipcam)) AS compdol, " _
            + vbCr + " 0 as reemdol,0 as lgddol,0 as letdol " _
            + vbCr + " FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN ((((((com_compras LEFT JOIN mae_documento ON com_compras.tipdoc = mae_documento.id) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) " _
            + vbCr + " LEFT JOIN var_ordendespacho ON com_compras.iddocref2 = var_ordendespacho.id) LEFT JOIN mae_cliente AS mae_cliente_1 ON var_ordendespacho.idcli = mae_cliente_1.id) LEFT JOIN mae_documento AS mae_documento_1 ON var_ordendespacho.idtipdoc = mae_documento_1.id) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro " _
            + vbCr + " WHERE (((com_compras.numreg)<>'000001')) " & nSQLFiltro
            '+ vbCr + " WHERE (((com_compras.fchreg) Between CDate('01/01/10') And CDate('31/01/10')) AND ((com_compras.numreg)<>'000001')) "
            
        '--fltro por fecha del documento
        If OptFch(0).Value = True Then nSQL = Replace(nSQL, "var_ordendespacho.fchemi between", "com_compras.fchdoc between")
        '--filtro por fecha de registro
        If OptFch(1).Value = True Then nSQL = Replace(nSQL, "var_ordendespacho.fchemi between", "com_compras.fchreg between")
        '--reporte sin igv
        If Option4.Value = True Then nSQL = Replace(nSQL, "[com_compras].[imptot]", "(com_compras.impbru + com_compras.impbru2 + com_compras.impbru3 + com_compras.impina)")
        
    End If
    
    '--Honorarios
    If NulosN(Fg3.TextMatrix(2, 2)) = -1 Then
        '--verificar si hay union de consultas
        If nSQL <> "" Then nSQL = nSQL + vbCr + " UNION "
        '--Los registros de honorarios se colocaran en la misma columna del registro de compras
        nSQL = nSQL _
            + vbCr + " SELECT 'Honorarios' as modulo, var_ordendespacho.id AS refid, mae_documento_1.abrev AS refabrev, mae_cliente_1.nombre AS refnomcli, var_ordendespacho.numerodoc AS refnumdoc, var_ordendespacho.fchemi AS reffchdoc, " _
            + vbCr + " com_honorarios.id AS iddoc, Left([com_honorarios].[numreg],2) & [mae_libros].[codsun] & Right([com_honorarios].[numreg],4) AS registro, mae_prov.numruc, mae_prov.nombre, mae_documento.abrev, [com_honorarios].[numser] & '-' & [com_honorarios].[numdoc] AS numdocumento, com_honorarios.fchdoc, com_honorarios.fchven, com_honorarios.glosa, mae_moneda.simbolo, " _
            + vbCr + " IIf([com_honorarios].[tc]=0,IIf([con_tc].[impven] Is Null,0,[con_tc].[impven]),[com_honorarios].[tc]) AS tipcam, IIf(com_honorarios.tipdoc=7,(-1)*[com_honorarios].[imptot],[com_honorarios].[imptot]) AS impreal, " _
            + vbCr + " 0 as vtasol, " _
            + vbCr + " IIf([com_honorarios].[idmon]=1,impreal,impreal*tipcam) AS compsol, " _
            + vbCr + " 0 as reemsol,0 as lgdsol,0 as letsol,  0 as vtadol,  " _
            + vbCr + " IIf([com_honorarios].[idmon]=2,impreal,IIf(tipcam=0,0,impreal/tipcam)) AS compdol, " _
            + vbCr + " 0 as reemdol,0 as lgddol,0 as letdol " _
            + vbCr + " FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN ((((((com_honorarios LEFT JOIN mae_documento ON com_honorarios.tipdoc = mae_documento.id) LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha) " _
            + vbCr + " LEFT JOIN var_ordendespacho ON com_honorarios.iddocref2 = var_ordendespacho.id) LEFT JOIN mae_cliente AS mae_cliente_1 ON var_ordendespacho.idcli = mae_cliente_1.id) LEFT JOIN mae_documento AS mae_documento_1 ON var_ordendespacho.idtipdoc = mae_documento_1.id) ON mae_moneda.id = com_honorarios.idmon) ON mae_prov.id = com_honorarios.idpro " _
            + vbCr + " WHERE ((com_honorarios.numreg)<>'000001') " & nSQLFiltro
            '+ vbCr + " WHERE (((com_honorarios.fchreg) Between CDate('01/01/10') And CDate('31/01/10')) AND ((com_honorarios.numreg)<>'000001')) "
            
        '--fltro por fecha del documento
        If OptFch(0).Value = True Then nSQL = Replace(nSQL, "var_ordendespacho.fchemi between", "com_honorarios.fchdoc between")
        '--filtro por fecha de registro
        If OptFch(1).Value = True Then nSQL = Replace(nSQL, "var_ordendespacho.fchemi between", "com_honorarios.fchreg between")
        '--reporte sin igv
        If Option4.Value = True Then nSQL = Replace(nSQL, "[com_honorarios].[imptot]", "(com_honorarios.impbru + com_honorarios.impina)")
            
    End If
    
    
    '--Reembolsables
    If NulosN(Fg3.TextMatrix(3, 2)) = -1 Then
        '--verificar si hay union de consultas
        If nSQL <> "" Then nSQL = nSQL + vbCr + " UNION "
        
        nSQL = nSQL _
            + vbCr + " SELECT 'Reembolsables' as modulo, var_ordendespacho.id AS refid, mae_documento_1.abrev AS refabrev, mae_cliente_1.nombre AS refnomcli, var_ordendespacho.numerodoc AS refnumdoc, var_ordendespacho.fchemi AS reffchdoc, " _
            + vbCr + " com_reembolsables.id AS iddoc, '' AS registro, mae_prov.numruc, mae_prov.nombre, mae_documento.abrev, [com_reembolsables].[numser] & '-' & [com_reembolsables].[numdoc] AS numdocumento, com_reembolsables.fchdoc, com_reembolsables.fchven, com_reembolsables.glosa, mae_moneda.simbolo, " _
            + vbCr + " IIf([com_reembolsables].[tc]=0,IIf([con_tc].[impven] Is Null,0,[con_tc].[impven]),[com_reembolsables].[tc]) AS tipcam, IIf(com_reembolsables.tipdoc=7,(-1)*[com_reembolsables].[imptot],[com_reembolsables].[imptot]) AS impreal, " _
            + vbCr + " 0 as vtasol, 0 as compsol, " _
            + vbCr + " IIf([com_reembolsables].[idmon]=1,impreal,impreal*tipcam) AS reemsol, " _
            + vbCr + " 0 as lgdsol,0 as letsol, 0 as vtadol, 0 as compdol, " _
            + vbCr + " IIf([com_reembolsables].[idmon]=2,impreal,IIf(tipcam=0,0,impreal/tipcam)) AS reemdol, " _
            + vbCr + " 0 as lgddol,0 as letdol " _
            + vbCr + " FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (((((com_reembolsables LEFT JOIN mae_documento ON com_reembolsables.tipdoc = mae_documento.id) LEFT JOIN con_tc ON com_reembolsables.fchdoc = con_tc.fecha) LEFT JOIN var_ordendespacho  " _
            + vbCr + " ON com_reembolsables.iddocref2 = var_ordendespacho.id) LEFT JOIN mae_cliente AS mae_cliente_1 ON var_ordendespacho.idcli = mae_cliente_1.id) LEFT JOIN mae_documento AS mae_documento_1 ON var_ordendespacho.idtipdoc = mae_documento_1.id) ON mae_moneda.id = com_reembolsables.idmon) ON mae_prov.id = com_reembolsables.idpro " _
            + vbCr + " WHERE ((com_reembolsables.idtipdocref) Is Not Null) " & nSQLFiltro
            '+ vbCr + " WHERE (((com_reembolsables.fchdoc) Between CDate('01/01/10') And CDate('31/01/10')) AND ((com_reembolsables.idtipdocref) Is Not Null)) "
            
        '--fltro por fecha del documento
        If OptFch(0).Value = True Then nSQL = Replace(nSQL, "var_ordendespacho.fchemi between", "com_reembolsables.fchdoc between")
        '--filtro por fecha de registro - No tiene fecha de registro; se considerara fecha de documento
        If OptFch(1).Value = True Then nSQL = Replace(nSQL, "var_ordendespacho.fchemi between", "com_reembolsables.fchdoc between")
        '--reporte sin igv
        If Option4.Value = True Then nSQL = Replace(nSQL, "[com_reembolsables].[imptot]", "(com_reembolsables.impbru + com_reembolsables.impina)")
            
            
    End If
    
    '--Liquidacion de Gasto Debito
    If NulosN(Fg3.TextMatrix(4, 2)) = -1 Then
        '--verificar si hay union de consultas
        If nSQL <> "" Then nSQL = nSQL + vbCr + " UNION "
        
        nSQL = nSQL _
            + vbCr + " SELECT 'Liq. Gasto Débito' as modulo, var_ordendespacho.id AS refid, mae_documento_1.abrev AS refabrev, mae_cliente_1.nombre AS refnomcli, var_ordendespacho.numerodoc AS refnumdoc, var_ordendespacho.fchemi AS reffchdoc, " _
            + vbCr + " vta_gastodebito.id AS iddoc, Left([vta_gastodebito].[numreg],2) & [mae_libros].[codsun] & Right([vta_gastodebito].[numreg],4) AS registro, mae_cliente.numruc, mae_cliente.nombre, mae_documento.abrev, [vta_gastodebito].[numser] & '-' & [vta_gastodebito].[numdoc] AS numdocumento, vta_gastodebito.fchemi as fchdoc, vta_gastodebito.fchven, vta_gastodebito.glosa, mae_moneda.simbolo, " _
            + vbCr + " IIf([vta_gastodebito].[tc]=0,IIf([con_tc].[impven] Is Null,0,[con_tc].[impven]),[vta_gastodebito].[tc]) AS tipcam,IIf([vta_gastodebito].[tipdoc]=7,(-1)*[vta_gastodebito].[imptot],[vta_gastodebito].[imptot]) AS impreal, " _
            + vbCr + " 0 as vtasol, 0 as compsol,0 as reemsol, " _
            + vbCr + " IIf([vta_gastodebito].[idmon]=1,impreal,impreal*tipcam) AS lgdsol, " _
            + vbCr + " 0 as letsol,  0 as vtadol, 0 as compdol,0 as reemdol, " _
            + vbCr + " IIf([vta_gastodebito].[idmon]=2,impreal,IIf(tipcam=0,0,impreal/tipcam)) AS lgddol, " _
            + vbCr + " 0 as letdol " _
            + vbCr + " FROM (((mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN (((vta_gastodebito LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id) LEFT JOIN mae_libros ON vta_gastodebito.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha) ON mae_moneda.id = vta_gastodebito.idmon) ON mae_cliente.id = vta_gastodebito.idcli) " _
            + vbCr + " LEFT JOIN var_ordendespacho ON vta_gastodebito.iddocref2 = var_ordendespacho.id) LEFT JOIN mae_cliente AS mae_cliente_1 ON var_ordendespacho.idcli = mae_cliente_1.id) LEFT JOIN mae_documento AS mae_documento_1 ON var_ordendespacho.idtipdoc = mae_documento_1.id " _
            + vbCr + " WHERE (((vta_gastodebito.numreg)<>'000001') AND ((vta_gastodebito.anulado)=0)) " & nSQLFiltro
            '+ vbCr + " WHERE (((vta_gastodebito.fchreg) Between CDate('01/01/10') And CDate('31/01/10')) AND ((vta_gastodebito.numreg)<>'000001') AND ((vta_gastodebito.anulado)=0)) "
            
        '--fltro por fecha del documento
        If OptFch(0).Value = True Then nSQL = Replace(nSQL, "var_ordendespacho.fchemi between", "vta_gastodebito.fchemi between")
        '--filtro por fecha de registro
        If OptFch(1).Value = True Then nSQL = Replace(nSQL, "var_ordendespacho.fchemi between", "vta_gastodebito.fchreg between")
        '--reporte sin igv
        If Option4.Value = True Then nSQL = Replace(nSQL, "[vta_gastodebito].[imptot]", "(vta_gastodebito.impbru + vta_gastodebito.impina)")
            
    End If
    
    
    '--Letras / Abonos
    If NulosN(Fg3.TextMatrix(5, 2)) = -1 Then
        '--verificar si hay union de consultas
        If nSQL <> "" Then nSQL = nSQL + vbCr + " UNION "
        
        nSQL = nSQL _
            + vbCr + " SELECT 'Letras' as modulo, var_ordendespacho.id AS refid, mae_documento_1.abrev AS refabrev, mae_cliente_1.nombre AS refnomcli, var_ordendespacho.numerodoc AS refnumdoc, var_ordendespacho.fchemi AS reffchdoc, " _
            + vbCr + " let_letradet.corr AS iddoc, Left([let_letra].[numreg],2) & [mae_libros].[codsun] & Right([let_letra].[numreg],4) AS registro, mae_cliente.numruc, mae_cliente.nombre, mae_documento.abrev, [let_letra].[ano] & ' ' & [let_letradet].[numdoc] & ' ' & [let_letradet].[numser] AS numdocumento, let_letradet.fchemi AS fchdoc, let_letradet.fchven, let_letra.glosa, mae_moneda.simbolo, " _
            + vbCr + " IIf([let_letra].[tc]=0,IIf([con_tc].[impven] Is Null,0,[con_tc].[impven]),[let_letra].[tc]) AS tipcam, let_letradet.implet AS impreal, " _
            + vbCr + " 0 as vtasol, 0 as compsol,0 as reemsol,0 as lgdsol, " _
            + vbCr + " IIf([let_letra].[idmon]=1,impreal,impreal*[tipcam]) AS letsol, " _
            + vbCr + " 0 as vtadol, 0 as compdol,0 as reemdol,0 as lgddol, " _
            + vbCr + " IIf([let_letra].[idmon] = 2, impreal, IIf([tipcam] = 0, 0, impreal / [tipcam])) AS letdol " _
            + vbCr + " FROM mae_moneda RIGHT JOIN (mae_libros RIGHT JOIN (mae_documento RIGHT JOIN ((((let_letra LEFT JOIN mae_cliente ON let_letra.idclipro = mae_cliente.id) LEFT JOIN ((var_ordendespacho LEFT JOIN mae_cliente AS mae_cliente_1 ON var_ordendespacho.idcli = mae_cliente_1.id) " _
            + vbCr + " LEFT JOIN mae_documento AS mae_documento_1 ON var_ordendespacho.idtipdoc = mae_documento_1.id) ON let_letra.iddocref2 = var_ordendespacho.id) LEFT JOIN con_tc ON let_letra.fchemi = con_tc.fecha) INNER JOIN let_letradet ON let_letra.id = let_letradet.idlet) ON mae_documento.id = let_letra.tipdoc) ON mae_libros.id = let_letra.idlib) ON mae_moneda.id = let_letra.idmon " _
            + vbCr + " WHERE (((let_letra.numreg)<>'000001')) " & nSQLFiltro
            '+ vbCr + " WHERE (((let_letra.fchreg) Between CDate('01/01/10') And CDate('31/01/10')) AND ((let_letra.numreg)<>'000001')) "
                
        '--fltro por fecha del documento
        If OptFch(0).Value = True Then nSQL = Replace(nSQL, "var_ordendespacho.fchemi between", "let_letra.fchemi between")
        '--filtro por fecha de registro
        If OptFch(1).Value = True Then nSQL = Replace(nSQL, "var_ordendespacho.fchemi between", "let_letra.fchreg between")
        '--reporte sin igv
        If Option4.Value = True Then nSQL = Replace(nSQL, "let_letradet.implet", "let_letradet.impcapital")
        
        '---------------------------------------------------------------------------------------------------------
        
        '--Abonos Bancarios (+)
        '--Cargos Bancarios (Devolucion a Clientes) (-) aplicable solo a operaciones cuya operacion es "CHEQUE GIRADO POR DEVOLUCION"
        
        nSQL = nSQL _
            + vbCr + " UNION " _
            + vbCr + " SELECT IIF(tes_caja.tipmov = 1,'Tesorería-Abonos','Tesorería-Cargos') as modulo, var_ordendespacho.id AS refid, mae_documento_1.abrev AS refabrev, mae_cliente_1.nombre AS refnomcli, var_ordendespacho.numerodoc AS refnumdoc, var_ordendespacho.fchemi AS reffchdoc, " _
            + vbCr + " tes_caja.id AS iddoc, Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4) AS registro, mae_bancos.numruc, mae_bancos.descripcion AS nombre, mae_documento.abrev, tes_cajaorigendet.numdoc AS numdocumento, tes_caja.fchope AS fchdoc, Null AS fchven, tes_caja.glosa, mae_moneda.simbolo, " _
            + vbCr + " IIf([tes_cajadestino].[tc]=0,IIf([con_tc].[impven] Is Null,0,[con_tc].[impven]),[tes_cajadestino].[tc]) AS tipcam, " _
            + vbCr + " IIF(tes_caja.tipmov = 1,1,-1) * ([tes_cajadestinodet].[importe]+[tes_cajadestinodet].[acuenta]) AS impreal, " _
            + vbCr + " 0 AS vtasol, 0 AS compsol, 0 AS reemsol, 0 AS lgdsol, " _
            + vbCr + " IIf([tes_caja].[idmon]=1,[impreal],[impreal]*[tipcam]) AS letsol, " _
            + vbCr + " 0 AS vtadol, 0 AS compdol, 0 AS reemdol, 0 AS lgddol, " _
            + vbCr + " IIf([tes_caja].[idmon] = 2, [impreal], IIf([tipcam] = 0, 0, [impreal] / [tipcam])) As letdol " _
            + vbCr + " FROM mae_bancos INNER JOIN ((((((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN tes_cajadestino ON tes_caja.id = tes_cajadestino.idtes) INNER JOIN (tes_cajadestinodet LEFT JOIN ((var_ordendespacho LEFT JOIN mae_cliente AS mae_cliente_1 ON var_ordendespacho.idcli = mae_cliente_1.id) " _
            + vbCr + " LEFT JOIN mae_documento AS mae_documento_1 ON var_ordendespacho.idtipdoc = mae_documento_1.id) ON tes_cajadestinodet.iddocref = var_ordendespacho.id) ON (tes_cajadestino.iddes = tes_cajadestinodet.iddes) AND (tes_cajadestino.idtes = tes_cajadestinodet.idtes)) INNER JOIN ((mae_banconumcta INNER JOIN tes_cajaori ON mae_banconumcta.id = tes_cajaori.idbcocta) " _
            + vbCr + " INNER JOIN ((tes_cajaorigendet LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id) LEFT JOIN mae_documento ON tes_documentos.iddoc = mae_documento.id) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes) ON mae_bancos.id = mae_banconumcta.idban " _
            + vbCr + " WHERE ( (((tes_caja.tipmov)=1) AND ((tes_cajadestinodet.iddocref)<>0) AND ((tes_cajaori.idbcocta)<>0)) " & nSQLFiltro & " ) OR " _
            + vbCr + "       ( (((tes_caja.tipmov)=2) AND ((tes_cajadestinodet.iddocref)<>0) AND ((tes_cajaori.idbcocta)<>0) AND ((tes_cajadestinodet.docctacte)='CHEQUE GIRADO POR DEVOLUCION')) " & nSQLFiltro & " )"

        
        '--Abonos Bancarios (Devoluciones) (-)
        '--Cargos Bancarios (Devoluciones a clientes)(+) aplicable solo a operaciones cuya operacion es "CHEQUE GIRADO POR DEVOLUCION"
        
        nSQL = nSQL _
            + vbCr + " UNION " _
            + vbCr + " SELECT IIF(tes_caja.tipmov = 1,'Tesorería-Abonos','Tesorería-Cargos') as modulo, var_ordendespacho.id AS refid, mae_documento_1.abrev AS refabrev, mae_cliente_1.nombre AS refnomcli, var_ordendespacho.numerodoc AS refnumdoc, var_ordendespacho.fchemi AS reffchdoc, " _
            + vbCr + " tes_caja.id AS iddoc, Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4) AS registro, mae_bancos.numruc, mae_bancos.descripcion AS nombre, mae_documento.abrev, tes_cajaorigendet.numdoc AS numdocumento, tes_caja.fchope AS fchdoc, Null AS fchven, tes_caja.glosa, mae_moneda.simbolo, " _
            + vbCr + " IIf([tes_cajaori].[tc]=0,IIf([con_tc].[impven] Is Null,0,[con_tc].[impven]),[tes_cajaori].[tc]) AS tipcam, " _
            + vbCr + " IIF(tes_caja.tipmov = 1,-1,1) * ([tes_cajaorigendet_1].[importe]+[tes_cajaorigendet_1].[acuenta]) AS impreal, " _
            + vbCr + " 0 AS vtasol, 0 AS compsol, 0 AS reemsol, 0 AS lgdsol, " _
            + vbCr + " IIf([tes_caja].[idmon]=1,[impreal],[impreal]*[tipcam]) AS letsol, " _
            + vbCr + " 0 AS vtadol, 0 AS compdol, 0 AS reemdol, 0 AS lgddol, " _
            + vbCr + " IIf([tes_caja].[idmon] = 2, [impreal], IIf([tipcam] = 0, 0, [impreal] / [tipcam])) As letdol " _
            + vbCr + " FROM ((mae_bancos INNER JOIN ((((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN con_tc ON tes_caja.fchope = con_tc.fecha) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) INNER JOIN ((mae_banconumcta INNER JOIN tes_cajaori ON mae_banconumcta.id = tes_cajaori.idbcocta) INNER JOIN ((tes_cajaorigendet LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id) " _
            + vbCr + " LEFT JOIN mae_documento ON tes_documentos.iddoc = mae_documento.id) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_caja.id = tes_cajaori.idtes) ON mae_bancos.id = mae_banconumcta.idban) INNER JOIN tes_cajaori AS tes_cajaori_1 ON tes_caja.id = tes_cajaori_1.idtes) INNER JOIN (tes_cajaorigendet AS tes_cajaorigendet_1 LEFT JOIN ((var_ordendespacho " _
            + vbCr + " LEFT JOIN mae_cliente AS mae_cliente_1 ON var_ordendespacho.idcli = mae_cliente_1.id) LEFT JOIN mae_documento AS mae_documento_1 ON var_ordendespacho.idtipdoc = mae_documento_1.id) ON tes_cajaorigendet_1.iddocref = var_ordendespacho.id) ON (tes_cajaori_1.idori = tes_cajaorigendet_1.idori) AND (tes_cajaori_1.idtes = tes_cajaorigendet_1.idtes) " _
            + vbCr + " WHERE ( (((tes_caja.tipmov)=1) AND ((tes_cajaori.idbcocta)<>0) AND ((tes_cajaori_1.idbcocta)=0) AND ((tes_cajaorigendet_1.iddocref)<>0)) " & nSQLFiltro & " ) OR " _
            + vbCr + "       ( (((tes_caja.tipmov)=2) AND ((tes_cajaori.idbcocta)<>0) AND ((tes_cajaori_1.idbcocta)=0) AND ((tes_cajaorigendet_1.iddocref)<>0) AND ((tes_cajaorigendet_1.docctacte)='CHEQUE GIRADO POR DEVOLUCION')) " & nSQLFiltro & " )"
                    
        
        '--fltro por fecha del documento
        If OptFch(0).Value = True Then nSQL = Replace(nSQL, "var_ordendespacho.fchemi between", "tes_caja.fchope between")
        '--filtro por fecha de registro
        If OptFch(1).Value = True Then nSQL = Replace(nSQL, "var_ordendespacho.fchemi between", "tes_caja.fchreg between")
        '--reporte sin igv
        'no hay filtro por este punto
        'If Option4.Value = True Then nSQL = Replace(nSQL, "let_letradet.implet", "let_letradet.impcapital")
    
    End If
           
    
    '--enviar la consulta
    GenerarConsulta = nSQL

End Function


Sub VerResumen()
    '===================================================================================================
    'creado: 24/09/10
    'Propósito: muestra consulta resumida en pantalla
    '
    'Entradas:  Indice = Ninguno
    '
    'Resultados: Consulta en pantalla segun parametros ingresados por usuario
    '
    'Modificado 27/12/10 Johan Castro
    '           Reiniciar variables que acumulan los totales.
    '===================================================================================================

    Dim Rst As New ADODB.Recordset
    Dim nSQL  As String
    Dim nSQLSub As String '--Sentencia SQL para identificar una subconsulta; está a nivel de detalle
    Dim nSQLSort As String '--Sentencia SQL para aplicar orden a la consulta
    Dim A As Double
    
    Dim xCliente As String '--Util para generar los grupos por cliente
    '----------
    '--Cuando cambia el grupo de cliente estas variables se reinican a valor =0
    Dim xTotDif1Sol As Double 'Acumular importes x Cliente de Diferencias Ventas-Compras
    Dim xTotDif2Sol As Double 'Acumular importes x Cliente de Diferencias LGD-Reemb
    Dim xTotDif3Sol  As Double 'Acumular importes x Cliente de Diferencias Anticipos -Facturacion
    '----------
    Dim xTotDif1SolTot As Double 'Acumular importes general de Diferencias Ventas-Compras
    Dim xTotDif2SolTot As Double 'Acumular importes general de Diferencias LGD-Reemb
    Dim xTotDif3SolTot  As Double 'Acumular importes general de Diferencias Anticipos -Facturacion
    
    '--Generar la sub consulta
    nSQLSub = GenerarConsulta()
    
        
    '--aplicar el orden
    If OptSort1.Value = True Then
        nSQLSort = ",det.reffchdoc "
    ElseIf OptSort2.Value = True Then
        nSQLSort = ",det.refnumdoc "
    ElseIf OptSort3.Value = True Then
        nSQLSort = ",det.registro "
    Else
        nSQLSort = ",det.reffchdoc,det.refnumdoc "
    End If
    
    '--reiniciar variables
    xTotVenTot = 0
    xTotComTot = 0
    xTotLGDTot = 0
    xTotReeTot = 0
    xTotLetTot = 0
    
    
    '--verificar la moneda a expresar la consulta
    If NulosN(TxtIdMon.Text) = 1 Then '--moneda nacional
        nSQL = "SELECT det.refnomcli,det.refabrev,det.refnumdoc,det.reffchdoc, " _
                + vbCr + " SUM(det.compsol) as imptotcomp ,SUM(det.vtasol) as imptotvta,SUM(det.reemsol) as imptotreem,SUM(det.lgdsol) as imptotlgd,SUM(det.letsol) as imptotlet, " _
                + vbCr + " imptotvta-imptotcomp as dif_vta_com, imptotlgd-imptotreem as dif_lgd_rem, (imptotlet-(imptotvta+imptotlgd)) as dif_ant_fact " _
                + vbCr + " FROM ( " _
                + vbCr + nSQLSub _
                + vbCr + " ) AS det GROUP BY det.refnomcli,det.refabrev,det.refnumdoc,det.reffchdoc " _
                + vbCr + " ORDER BY det.refnomcli, det.refnumdoc, det.reffchdoc "
                
    Else '--moneda extranjera
        nSQL = "SELECT det.refnomcli,det.refabrev,det.refnumdoc,det.reffchdoc, " _
                + vbCr + " SUM(det.compdol) as imptotcomp ,SUM(det.vtadol) as imptotvta,SUM(det.reemdol) as imptotreem,SUM(det.lgddol) as imptotlgd,SUM(det.letdol) as imptotlet, " _
                + vbCr + " imptotvta-imptotcomp as dif_vta_com, imptotlgd-imptotreem as dif_lgd_rem, (imptotlet-(imptotvta+imptotlgd)) as dif_ant_fact " _
                + vbCr + " FROM ( " _
                + vbCr + nSQLSub _
                + vbCr + " ) AS det GROUP BY det.refnomcli,det.refabrev,det.refnumdoc,det.reffchdoc " _
                + vbCr + " ORDER BY det.refnomcli " & nSQLSort
    End If
    
    '--cambiar cursor de espera del mouse
    Me.MousePointer = vbHourglass
    
    '--ejecutar la consulta
    RST_Busq Rst, nSQL, xCon
    
    Fg1.Rows = 2
    If Rst.RecordCount <> 0 Then
            
        '--centrar la barra de progreso
        Frame5.Left = (Me.Width - Frame5.Width) / 2
        Frame5.Top = (Me.Height - Frame5.Height) / 2
        '--obtener cantidad de registros
        ProgressBar1.Max = Rst.RecordCount
        '--mostrar la barra de progreso
        Frame5.Visible = True
        
        
        Rst.MoveFirst
                
        For A = 1 To Rst.RecordCount
            
            ProgressBar1.Value = A
            
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(Rst("refnomcli"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(Rst("refabrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(Rst("refnumdoc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(NulosC(Rst("reffchdoc")), "dd/mm/yy")
            If IsNull(Rst("reffchdoc")) = False Then Fg1.TextMatrix(Fg1.Rows - 1, 5) = (CDate(TxtFchFin.Valor) - Rst("reffchdoc"))
            '--
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(Rst("imptotcomp"), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(Rst("imptotvta"), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(Rst("imptotreem"), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(Rst("imptotlgd"), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(Rst("imptotlet"), FORMAT_MONTO)
            '--diferencias
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(Rst("dif_vta_com"), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(Rst("dif_lgd_rem"), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 13) = Format(Rst("dif_ant_fact"), FORMAT_MONTO)
                        
            '--acumulando para mostrar resumen por orden
            xTotCom = xTotCom + NulosN(Rst("imptotcomp"))
            xTotVen = xTotVen + NulosN(Rst("imptotvta"))
            xTotRee = xTotRee + NulosN(Rst("imptotreem"))
            xTotLGD = xTotLGD + NulosN(Rst("imptotlgd"))
            xTotLet = xTotLet + NulosN(Rst("imptotlet"))
            
            xTotDif1Sol = xTotDif1Sol + NulosN(Rst("dif_vta_com"))
            xTotDif2Sol = xTotDif2Sol + NulosN(Rst("dif_lgd_rem"))
            xTotDif3Sol = xTotDif3Sol + NulosN(Rst("dif_ant_fact"))
            
            '------------------------------------------------------------------------
            '--acumulando los subtotales
            xTotComTot = xTotComTot + NulosN(Rst("imptotcomp"))
            xTotVenTot = xTotVenTot + NulosN(Rst("imptotvta"))
            xTotReeTot = xTotReeTot + NulosN(Rst("imptotreem"))
            xTotLGDTot = xTotLGDTot + NulosN(Rst("imptotlgd"))
            xTotLetTot = xTotLetTot + NulosN(Rst("imptotlet"))
            
            xTotDif1SolTot = xTotDif1SolTot + NulosN(Rst("dif_vta_com"))
            xTotDif2SolTot = xTotDif2SolTot + NulosN(Rst("dif_lgd_rem"))
            xTotDif3SolTot = xTotDif3SolTot + NulosN(Rst("dif_ant_fact"))
                        
            xCliente = NulosC(Rst("refnomcli"))
                        
            Rst.MoveNext
            If Rst.EOF = True Then
                
                Fg1.Rows = Fg1.Rows + 1
                
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, &H800080, True, &HE2FEFB, "TOTAL CLIENTE ==> "
                
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 6, &H80000012, True, &HE2FEFB, Format(xTotCom, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, &H80000012, True, &HE2FEFB, Format(xTotVen, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 8, &H80000012, True, &HE2FEFB, Format(xTotRee, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, &H80000012, True, &HE2FEFB, Format(xTotLGD, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, &H80000012, True, &HE2FEFB, Format(xTotLet, FORMAT_MONTO)
                
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H80000012, True, &HE2FEFB, Format(xTotDif1Sol, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H80000012, True, &HE2FEFB, Format(xTotDif2Sol, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 13, &H80000012, True, &HE2FEFB, Format(xTotDif3Sol, FORMAT_MONTO)
            
                Fg1.Rows = Fg1.Rows + 1
                
                xTotCom = 0
                xTotVen = 0
                xTotRee = 0
                xTotLGD = 0
                xTotLet = 0
                
                xTotDif1Sol = 0
                xTotDif2Sol = 0
                xTotDif3Sol = 0
                
                Exit For
            End If
            
            If xCliente <> NulosC(Rst("refnomcli")) Then
                Fg1.Rows = Fg1.Rows + 1
                
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, &H800080, True, &HE2FEFB, "TOTAL CLIENTE ==> "
                
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 6, &H80000012, True, &HE2FEFB, Format(xTotCom, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, &H80000012, True, &HE2FEFB, Format(xTotVen, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 8, &H80000012, True, &HE2FEFB, Format(xTotRee, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, &H80000012, True, &HE2FEFB, Format(xTotLGD, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, &H80000012, True, &HE2FEFB, Format(xTotLet, FORMAT_MONTO)
                
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H80000012, True, &HE2FEFB, Format(xTotDif1Sol, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H80000012, True, &HE2FEFB, Format(xTotDif2Sol, FORMAT_MONTO)
                FORMATO_CELDA Fg1, Fg1.Rows - 1, 13, &H80000012, True, &HE2FEFB, Format(xTotDif3Sol, FORMAT_MONTO)
            
                Fg1.Rows = Fg1.Rows + 1
                xTotCom = 0
                xTotVen = 0
                xTotRee = 0
                xTotLGD = 0
                xTotLet = 0
                
                xTotDif1Sol = 0
                xTotDif2Sol = 0
                xTotDif3Sol = 0
                
                xCliente = NulosC(Rst("refnomcli"))
            End If
        
        Next A
        
        Fg1.Rows = Fg1.Rows + 1
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 3, &H800080, True, &HE2FEFB, "GRAN TOTAL  ==> "
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 6, &H80000012, True, &HE2FEFB, Format(xTotComTot, FORMAT_MONTO)
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 7, &H80000012, True, &HE2FEFB, Format(xTotVenTot, FORMAT_MONTO)
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 8, &H80000012, True, &HE2FEFB, Format(xTotReeTot, FORMAT_MONTO)
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 9, &H80000012, True, &HE2FEFB, Format(xTotLGDTot, FORMAT_MONTO)
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 10, &H80000012, True, &HE2FEFB, Format(xTotLetTot, FORMAT_MONTO)
        
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 11, &H80000012, True, &HE2FEFB, Format(xTotDif1SolTot, FORMAT_MONTO)
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 12, &H80000012, True, &HE2FEFB, Format(xTotDif2SolTot, FORMAT_MONTO)
        FORMATO_CELDA Fg1, Fg1.Rows - 1, 13, &H80000012, True, &HE2FEFB, Format(xTotDif3SolTot, FORMAT_MONTO)
        
    End If
    
LaCague:
    Frame5.Visible = False
    Set Rst = Nothing
    Me.MousePointer = vbDefault
End Sub


Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 3 Then '--exportar
        
        If ButtonMenu.Index = 1 Then '--formato rapido
            pExportar True
        End If
        
        
        If ButtonMenu.Index = 2 Then '--formato lento
            pExportar False
        End If
    End If
  

End Sub


