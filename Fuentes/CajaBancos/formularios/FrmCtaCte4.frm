VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCtaCte4 
   Caption         =   "Caja y Bancos - Analisis"
   ClientHeight    =   7800
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   14280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDetalle 
      BorderStyle     =   0  'None
      Caption         =   "[ Depurar Datos ]"
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
      Height          =   4305
      Left            =   10050
      TabIndex        =   45
      Top             =   2580
      Visible         =   0   'False
      Width           =   10755
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Exportar MsExcel"
         Height          =   375
         Left            =   8970
         TabIndex        =   48
         Top             =   3930
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   3990
         TabIndex        =   47
         Top             =   3840
         Width           =   1755
      End
      Begin VB.PictureBox pic1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   10440
         Picture         =   "FrmCtaCte4.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   46
         ToolTipText     =   "Cerrar"
         Top             =   60
         Width           =   195
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg3 
         Height          =   2895
         Left            =   30
         TabIndex        =   49
         Top             =   870
         Width           =   10635
         _cx             =   18759
         _cy             =   5106
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
         BackColor       =   14745342
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   8388608
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   14745342
         GridColor       =   -2147483627
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCtaCte4.frx":02EC
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
      Begin VB.Label lblDetTC1 
         AutoSize        =   -1  'True
         Caption         =   "lblDetTC1"
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
         Left            =   8400
         TabIndex        =   66
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "T.C:"
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
         Left            =   7980
         TabIndex        =   65
         Top             =   660
         Width           =   375
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Emi."
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
         Left            =   3480
         TabIndex        =   64
         Top             =   660
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Importe:"
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
         Left            =   5670
         TabIndex        =   63
         Top             =   660
         Width           =   705
      End
      Begin VB.Label lblDetFchEmi1 
         AutoSize        =   -1  'True
         Caption         =   "lblDetFchEmi1"
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
         Left            =   4380
         TabIndex        =   62
         Top             =   660
         Width           =   1230
      End
      Begin VB.Label lblDetImp1 
         AutoSize        =   -1  'True
         Caption         =   "lblDetImp1"
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
         Left            =   6510
         TabIndex        =   61
         Top             =   660
         Width           =   1275
      End
      Begin VB.Label lblDetNumDoc1 
         Caption         =   "lblDetNumDoc1"
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
         Left            =   1230
         TabIndex        =   60
         Top             =   660
         Width           =   2205
      End
      Begin VB.Label lblDetNombre1 
         AutoSize        =   -1  'True
         Caption         =   "lblDetNombre1"
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
         Left            =   1230
         TabIndex        =   59
         Top             =   390
         Width           =   6900
      End
      Begin VB.Label lblDetNumDoc 
         AutoSize        =   -1  'True
         Caption         =   "Documento:"
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
         Left            =   60
         TabIndex        =   58
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label lblDetNombre 
         AutoSize        =   -1  'True
         Caption         =   "Nombres:"
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
         Left            =   60
         TabIndex        =   57
         Top             =   390
         Width           =   810
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de Documentos "
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
         TabIndex        =   50
         Top             =   90
         Width           =   2055
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   5
         X1              =   -30
         X2              =   10770
         Y1              =   4290
         Y2              =   4290
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   4
         X1              =   10740
         X2              =   10740
         Y1              =   0
         Y2              =   6970
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   5
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   6500
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   6
         X1              =   -30
         X2              =   11970
         Y1              =   0
         Y2              =   15
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   420
         Index           =   1
         Left            =   -690
         Top             =   -90
         Width           =   11370
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne2 
      Height          =   1545
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   13770
      _cx             =   24289
      _cy             =   2725
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
      CurrTab         =   0
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
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   1455
         Left            =   345
         TabIndex        =   7
         Top             =   45
         Width           =   13380
         Begin VB.Frame FraReem 
            Caption         =   "[ Seleccionar "
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
            Height          =   1005
            Left            =   11460
            TabIndex        =   51
            Top             =   -30
            Visible         =   0   'False
            Width           =   1575
            Begin VB.OptionButton OptReem2 
               Caption         =   "Lgd"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   210
               TabIndex        =   53
               Top             =   750
               Width           =   690
            End
            Begin VB.OptionButton OptReem1 
               Caption         =   "Bancos"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   210
               TabIndex        =   52
               Top             =   480
               Value           =   -1  'True
               Width           =   960
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Reembolsables ]"
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
               Left            =   90
               TabIndex        =   56
               Top             =   180
               Width           =   1410
            End
         End
         Begin VB.Frame Frame5 
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
            Height          =   930
            Left            =   1560
            TabIndex        =   54
            Top             =   -30
            Width           =   2505
            Begin VSFlex7Ctl.VSFlexGrid Fg4 
               Height          =   765
               Left            =   60
               TabIndex        =   55
               Top             =   165
               Width           =   2325
               _cx             =   4101
               _cy             =   1349
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   8388608
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   16777215
               GridColor       =   -2147483627
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
               Rows            =   10
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmCtaCte4.frx":0458
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
         Begin VB.Frame Frame1 
            Caption         =   "[Seleccionar Fecha]"
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
            TabIndex        =   20
            Top             =   870
            Width           =   4035
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
               Height          =   300
               Left            =   705
               TabIndex        =   21
               Top             =   225
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
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
               Height          =   300
               Left            =   2670
               TabIndex        =   22
               Top             =   225
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
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Hasta"
               Height          =   195
               Index           =   2
               Left            =   2145
               TabIndex        =   24
               Top             =   330
               Width           =   420
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Desde"
               Height          =   195
               Index           =   1
               Left            =   105
               TabIndex        =   23
               Top             =   330
               Width           =   465
            End
         End
         Begin VB.CheckBox Chk 
            Caption         =   "Ver Lineal"
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
            Height          =   435
            Left            =   8280
            TabIndex        =   44
            Top             =   960
            Width           =   855
         End
         Begin VB.Frame Frame3 
            Caption         =   "[  Seleccionar  ]"
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
            Left            =   4110
            TabIndex        =   8
            Top             =   -30
            Width           =   5100
            Begin VB.CommandButton CmdBusCliPro 
               Enabled         =   0   'False
               Height          =   240
               Left            =   4770
               Picture         =   "FrmCtaCte4.frx":0547
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   480
               Width           =   210
            End
            Begin VB.OptionButton OptSel2 
               Caption         =   "Seleccionar"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1500
               TabIndex        =   10
               Top             =   240
               Width           =   1140
            End
            Begin VB.OptionButton OptSel1 
               Caption         =   "Todos"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   150
               TabIndex        =   9
               Top             =   240
               Value           =   -1  'True
               Width           =   840
            End
            Begin VB.TextBox TxtCliPro 
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   300
               Left            =   165
               Locked          =   -1  'True
               TabIndex        =   26
               Text            =   "TxtCliPro"
               Top             =   450
               Width           =   4845
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Cliente"
               Height          =   195
               Index           =   0
               Left            =   2580
               TabIndex        =   28
               Top             =   210
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Label LblIdCliPro 
               Caption         =   "LblIdCliPro"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   3090
               TabIndex        =   27
               Top             =   240
               Visible         =   0   'False
               Width           =   750
            End
         End
         Begin VB.Frame Frame12 
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
            TabIndex        =   33
            Top             =   -30
            Width           =   1515
            Begin VB.OptionButton OptFch 
               Caption         =   "Fch Reg"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   35
               Top             =   435
               Value           =   -1  'True
               Width           =   1125
            End
            Begin VB.OptionButton OptFch 
               Caption         =   "Fch Doc"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   34
               Top             =   210
               Width           =   1140
            End
         End
         Begin VB.Frame Frame6 
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
            Height          =   585
            Left            =   4110
            TabIndex        =   15
            Top             =   870
            Width           =   4035
            Begin VB.CommandButton CmdBusMon 
               Height          =   240
               Left            =   1185
               Picture         =   "FrmCtaCte4.frx":0679
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   270
               Width           =   210
            End
            Begin VB.TextBox TxtIdMon 
               Height          =   300
               Left            =   720
               MaxLength       =   1
               TabIndex        =   17
               Text            =   "TxtIdMon"
               Top             =   240
               Width           =   705
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
               Left            =   1425
               TabIndex        =   19
               Top             =   240
               Width           =   2490
            End
            Begin VB.Label LblTipCam 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Moneda"
               Height          =   195
               Index           =   4
               Left            =   90
               TabIndex        =   18
               Top             =   330
               Width           =   585
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "[  Seleccionar Estado ]"
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
            Height          =   1485
            Left            =   9210
            TabIndex        =   29
            Top             =   -30
            Width           =   2220
            Begin VB.CheckBox chk_descuadrado 
               Caption         =   "Descuadrados"
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
               Left            =   150
               TabIndex        =   36
               ToolTipText     =   "Mostrará solo los documentos cuyo saldo final es negativo"
               Top             =   1140
               Width           =   1545
            End
            Begin VB.OptionButton OptPen 
               Caption         =   "Pendientes"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   150
               TabIndex        =   32
               Top             =   240
               Value           =   -1  'True
               Width           =   1110
            End
            Begin VB.OptionButton OptCan 
               Caption         =   "Cancelados"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   150
               TabIndex        =   31
               Top             =   495
               Width           =   1350
            End
            Begin VB.OptionButton OptTodos 
               Caption         =   "Todos"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   150
               TabIndex        =   30
               Top             =   750
               Width           =   900
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00FFFFFF&
               Index           =   0
               X1              =   120
               X2              =   2070
               Y1              =   1040
               Y2              =   1040
            End
            Begin VB.Line Line1 
               X1              =   120
               X2              =   2070
               Y1              =   1020
               Y2              =   1020
            End
         End
      End
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   14715
         TabIndex        =   6
         Top             =   45
         Width           =   13380
         Begin VB.Frame Fra_Orden 
            Caption         =   "[Aplicar Orden]"
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
            Height          =   1035
            Left            =   150
            TabIndex        =   37
            Top             =   90
            Width           =   1725
            Begin VB.OptionButton Opt_Orden 
               Caption         =   "N° Documento"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   40
               Top             =   270
               Width           =   1395
            End
            Begin VB.OptionButton Opt_Orden 
               Caption         =   "N°. Registro"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   39
               Top             =   510
               Width           =   1395
            End
            Begin VB.OptionButton Opt_Orden 
               Caption         =   "Fecha Doc."
               Height          =   195
               Index           =   2
               Left            =   60
               TabIndex        =   38
               Top             =   750
               Width           =   1395
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "[ Documentos de Apertura ]"
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
            Height          =   1035
            Left            =   2010
            TabIndex        =   11
            Top             =   90
            Width           =   2670
            Begin VB.OptionButton OptAperturaSolo 
               Caption         =   "Ver solo Apertura"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   225
               TabIndex        =   14
               Top             =   750
               Width           =   2070
            End
            Begin VB.OptionButton OptAperturaSin 
               Caption         =   "No incluir Apertura"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   225
               TabIndex        =   13
               Top             =   510
               Width           =   1830
            End
            Begin VB.OptionButton OptAperturaCon 
               Caption         =   "Incluir Apertura"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   225
               TabIndex        =   12
               Top             =   270
               Value           =   -1  'True
               Width           =   1710
            End
         End
      End
   End
   Begin VB.Frame fraBarra 
      BorderStyle     =   0  'None
      Caption         =   "FrmConsultaDiario"
      Height          =   780
      Left            =   60
      TabIndex        =   1
      Top             =   7650
      Visible         =   0   'False
      Width           =   6285
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   330
         Left            =   150
         TabIndex        =   2
         Top             =   315
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   582
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   -15
         Y2              =   900
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   6270
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   1
         X1              =   -75
         X2              =   6500
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   6270
         X2              =   6270
         Y1              =   -30
         Y2              =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Procesando Documentos"
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
         Left            =   195
         TabIndex        =   4
         Top             =   75
         Width           =   2130
      End
      Begin VB.Label lbl 
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
         Left            =   4605
         TabIndex        =   3
         Top             =   75
         Width           =   1530
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14280
      _ExtentX        =   25188
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
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3285
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
               Picture         =   "FrmCtaCte4.frx":07AB
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte4.frx":0CEF
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte4.frx":1081
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte4.frx":11DB
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte4.frx":156D
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte4.frx":16F1
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte4.frx":1B45
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte4.frx":1C5D
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte4.frx":21A1
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte4.frx":26E5
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte4.frx":27F9
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte4.frx":290D
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte4.frx":2D61
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmCtaCte4.frx":2ECD
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   5655
      Left            =   0
      TabIndex        =   41
      Top             =   1920
      Width           =   11880
      _cx             =   20955
      _cy             =   9975
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
      Appearance      =   2
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   14215660
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "      Detalle     |      Resumen     "
      Align           =   0
      CurrTab         =   1
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
      Begin VSFlex7Ctl.VSFlexGrid Fg1 
         Height          =   5235
         Left            =   -12435
         TabIndex        =   42
         Top             =   45
         Width           =   11790
         _cx             =   20796
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   8388608
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
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
         Rows            =   2
         Cols            =   13
         FixedRows       =   2
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCtaCte4.frx":3415
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
         Height          =   5235
         Left            =   45
         TabIndex        =   43
         Top             =   45
         Width           =   11790
         _cx             =   20796
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
         BackColorSel    =   8388608
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
         Rows            =   1
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCtaCte4.frx":359C
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
Attribute VB_Name = "FrmCtaCte4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--modificado 30/12/09 por Johan Castro
'       considerar los canjes de las notas de debito en letras(idlib=37)
'--modificado 13/01/10 por Johan Castro
'       rpt Ventas:mostrar NC cuando no esten vinculado a un documento de referencia
'--modificado 19/01/10 por Johan Castro
'       considerar el filtro por numdoc,registro,fecha
'--modificado 11/02/10 por Johan Castro
'       considerar ajuste por diferencia de cambio a proveedores,honorarios,cliente
'--modificado 14/05/10 por Johan Castro
'       considerar filtro por intervalo de fechas
'       considerar filtro por documentos de apertura
'--modificado 21/05/10 por Johan Castro
'       no mostrar las nc de ventas cuando son anulados
'--Modificado: 15/06/11 Johan Castro
'--     Dar flexibilidad al formulario para definir el tamaño del mismo

Option Explicit
Dim RstCta As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE
Dim xIdLibro As Integer     '--Codigo del libro
Dim xIdLibroRef As String   '--Codigo de libros que hacen referencia a libro xIdLibro
Dim OrigFX As Long '--para mover el frame posicion horizontal
Dim OrigFY As Long '--para mover el frame posicion vertical
Dim xRstTot As New ADODB.Recordset '--Rst para acumular los totales por cliente, proveedor en rpt lineal

Private Sub Chk_Click()
    pConfigurarGrilla
End Sub

Private Sub CmdBusCliPro_Click()
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
    

'''    If OptCliente.Value = True Then
'''        xForm.Titulo = "Buscando Clientes"
'''        xForm.SQLCad = "SELECT mae_cliente.numruc, mae_cliente.nombre, mae_cliente.id From mae_cliente ORDER BY mae_cliente.nombre"
'''        xCampos(0, 0) = "Cliente":       xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
'''    ElseIf OptProvee.Value = True Then
'''        xForm.Titulo = "Buscando Proveedores"
'''        xForm.SQLCad = "SELECT mae_prov.numruc, mae_prov.nombre, mae_prov.id From mae_prov ORDER BY mae_prov.nombre"
'''        xCampos(0, 0) = "Proveedor":     xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
'''    ElseIf opt4ta.Value = True Then
'''        xForm.Titulo = "Buscando Prestador de Servicio"
'''        xForm.SQLCad = "SELECT mae_prov.numruc, mae_prov.nombre, mae_prov.id From mae_prov WHERE mae_prov.tipper = 1 ORDER BY mae_prov.nombre"
'''        xCampos(0, 0) = "Proveedor":     xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
'''    End If

   '-----------------------------------------------
    If Fg4.Row < 1 Then Exit Sub
    
    xCampos(0, 0) = "Cliente":       xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"


    Select Case NulosN(Fg4.TextMatrix(Fg4.Row, 2))
        Case 1, 999, 4
            xForm.Titulo = "Buscando Proveedores"
            xForm.SQLCad = "SELECT mae_prov.numruc, mae_prov.nombre, mae_prov.id From mae_prov Where mae_prov.id <>0 ORDER BY mae_prov.nombre"
            xCampos(0, 0) = "Proveedor"
        Case 40
            xForm.Titulo = "Buscando Prestador de Servicio"
            xForm.SQLCad = "SELECT mae_prov.numruc, mae_prov.nombre, mae_prov.id From mae_prov WHERE mae_prov.tipper = 1 and mae_prov.id <>0 ORDER BY mae_prov.nombre"
            xCampos(0, 0) = "Proveedor"
        Case 9
            xForm.Titulo = "Buscando Empleados"
            xForm.SQLCad = "SELECT pla_empleados.numdoc as numruc, pla_empleados.nombre, pla_empleados.id From pla_empleados Where pla_empleados.id<>0  ORDER BY pla_empleados.nombre"
            xCampos(0, 0) = "Empleado"
        Case 2, 37, 41
            xForm.Titulo = "Buscando Clientes"
            xForm.SQLCad = "SELECT mae_cliente.numruc, mae_cliente.nombre, mae_cliente.id From mae_cliente Where mae_cliente.id<>0 ORDER BY mae_cliente.nombre"
            
        Case 42
            xForm.Titulo = "Buscando Banco"
            xForm.SQLCad = "SELECT mae_bancos.numruc, mae_bancos.descripcion as nombre, mae_bancos.id From mae_bancos  where mae_bancos.id <>0 ORDER BY mae_bancos.descripcion"
            xCampos(0, 0) = "Banco"
        
    End Select

    xCampos(1, 0) = "Nº R.U.C.":     xCampos(1, 1) = "numruc":        xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"

    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "nombre"
    xForm.CampoBusca = "nombre"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtCliPro.Text = NulosC(xRs("nombre"))
        LblIdCliPro.Caption = xRs("id")
        TxtFchIni.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub Fg1_DblClick()
    If Chk.Value = 1 And Fg1.Rows > Fg1.FixedRows Then VerDatosDetalle
End Sub

Private Sub Fg4_RowColChange()
    TxtCliPro.Text = ""
    LblIdCliPro.Caption = ""
    FraReem.Visible = False
    If NulosN(Fg4.TextMatrix(Fg4.Row, 2)) = 999 Then
        FraReem.Visible = True
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        pConfigurarGrilla
        TxtIdMon.Text = 1
        TxtIdMon_Validate False
        
        TabOne2.CurrTab = 0
        
        SeEjecuto = True
        
        '--colocar la fecha del primer dia del año de trabajo
        TxtFchIni.Valor = CDate("01/01/" & AnoTra)
        
        '--verificar si el año de trabajo es igual al año actual
        If NulosC(Year(Date)) < AnoTra Then
            TxtFchFin.Valor = CDate("31/12/" & AnoTra)
        Else
            TxtFchFin.Valor = Date
        End If
        
        '--enfocar el cursor en la fecha inicial
        TxtFchIni.SetFocus
    End If
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    ElseIf KeyCode = vbKeyF3 And Shift = 0 Then
        BuscarVSFlexGrid
    ElseIf KeyCode = vbKeyF8 Then
        pConsultar
    End If

End Sub

Private Sub Form_Load()
    TxtCliPro.Text = ""
    TxtFchIni.Valor = ""
    TxtFchIni.Valor = Date
    LblMoneda.Caption = ""
    TxtIdMon.Text = ""
    SeEjecuto = False

    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
  
    If Me.Height > 3000 Then
        TabOne1.Top = 1920
        TabOne1.Width = Me.Width - 150
        TabOne1.Height = Me.Height - 2320
    End If
End Sub



Private Sub OptCan_Click()
'    chk_descuadrado.Enabled = True
End Sub



Private Sub OptSel1_Click()
    TxtCliPro.Text = ""
    LblIdCliPro.Caption = ""
    TxtCliPro.Enabled = False
    CmdBusCliPro.Enabled = False
End Sub

Private Sub OptSel2_Click()
    TxtCliPro.Enabled = True
    CmdBusCliPro.Enabled = True
End Sub

Private Sub OptTodos_Click()
'    chk_descuadrado.Enabled = True
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 3 Then pExportar
    If Button.Index = 4 Then pImprimir
    If Button.Index = 6 Then
        Set RstCta = Nothing
        Unload Me
    End If
End Sub

Sub pExportar()
    On Error GoTo error
    Dim oExport As New SGI2_funciones.formularios
    Dim nPeriodo As String
    Dim nTitulo1 As String
    
    nPeriodo = "Al  " + CStr(TxtFchIni.Valor)

    nTitulo1 = "(Expresado en " & LblMoneda.Caption & ")"
    
    If TabOne1.CurrTab = 0 Then '--detalle
        '''oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "Cuenta Corriente - " + IIf(OptCliente.Value = True, "Cliente", "Proveedor"), nPeriodo, nTitulo1, "Cuenta Corriente Análisis"
        
        GRID_EXPORTAR_MSEXCELTMP Fg1, xCon, flexFileCustomText, True, "Cuenta Corriente - " + Fg3.TextMatrix(Fg3.Row, 1), "Expresado en " & LblMoneda.Caption, "Cuenta Corriente Análisis"

    Else
        oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg2, "Resumen de Cuenta Corriente - " + Fg3.TextMatrix(Fg3.Row, 1), nPeriodo, nTitulo1, "Cuenta Corriente Análisis"
    End If
    
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportar"
End Sub

Private Sub TxtCliPro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCliPro_Click
    End If
End Sub

'***********************************************************************************************
'------------CAMBIOS AL 020108

Private Sub CmdBusMon_Click()
    
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre":      xCampos(0, 1) = "descripcion":     xCampos(0, 2) = "4500":      xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":   xCampos(1, 1) = "id":              xCampos(1, 2) = "500":      xCampos(1, 3) = "N"

    CARGAR_DLL_EPSBUSCAR xCon, xRs, "SELECT * FROM mae_moneda ORDER BY descripcion ;", xCampos(), "Buscando Moneda", "descripcion", "descripcion", Principio
    
    If xRs.State = 0 Then GoTo Salir
    If xRs.RecordCount = 0 Then GoTo Salir
    TxtIdMon.Text = xRs("id") & ""
    LblMoneda.Caption = xRs("descripcion") & ""
    
Salir:
    Set xRs = Nothing
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

Private Sub BuscarVSFlexGrid()
    On Error GoTo error
    
    Dim oExport As New SGI2_funciones.formularios
    Dim xCampos(4, 3) As String
    'campo     'columna del grid    'tipo(N:Numerico, C:caracter, F:fecha)      campo predeterminado(0:no se muestra, -1:se muestra al iniciar el formulario)
    xCampos(0, 0) = "Nº.Registro":      xCampos(0, 1) = "1":    xCampos(0, 2) = "C":    xCampos(0, 3) = "-1"
    xCampos(1, 0) = "Origen":           xCampos(1, 1) = "2":    xCampos(1, 2) = "C":    xCampos(1, 3) = "0"
    xCampos(2, 0) = "Nº Documento":     xCampos(2, 1) = "4":    xCampos(2, 2) = "C":    xCampos(2, 3) = "0"
    xCampos(3, 0) = Label1(0):          xCampos(3, 1) = "4":    xCampos(3, 2) = "C":    xCampos(3, 3) = "0"
    xCampos(4, 0) = "Fch.Emi.":         xCampos(4, 1) = "5":    xCampos(4, 2) = "F":    xCampos(4, 3) = "0"
    
    oExport.VSFlexGrid_Buscar Me.hWnd, Fg1, xCampos(), Fg1.Row
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "BuscarVSFlexGrid"
End Sub

Private Sub pConfigurarGrilla()
    
    '--PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA
    Dim A As Integer
    TabOne1.CurrTab = 0
    
    Fg1.Rows = 1
    Fg1.Cols = 1
    Fg2.Rows = 1
    Fg2.Cols = 1
    
    Fg1.ExplorerBar = flexExNone
    Fg2.ExplorerBar = flexExNone
    
    Fg1.AutoSearch = flexSearchFromTop
    Fg2.AutoSearch = flexSearchFromTop
    
    If Chk.Value = 0 Then
        With Fg1
            '-----
            .Rows = 2
            .Cols = 17
            .FixedRows = 2
            .FrozenCols = 0
            .RowHeight(0) = 250
            .ColWidth(0) = 200
            UNIR_CELDAS Fg1, 0, 1, 0, 9, "DATOS DEL DOCUMENTO", flexAlignCenterCenter
            FORMATO_CELDA Fg1, 0, 1, vbBlack, True, &HD8E9EC
            If Trim(LblMoneda.Caption) = "" Then
                UNIR_CELDAS Fg1, 0, 10, 0, 12, "IMPORTES", flexAlignCenterCenter
            Else
                UNIR_CELDAS Fg1, 0, 10, 0, 12, "IMPORTES EN " & UCase(LblMoneda.Caption), flexAlignCenterCenter
            End If
            FORMATO_CELDA Fg1, 0, 10, vbBlack, True, &HD8E9EC
            
            UNIR_CELDAS Fg1, 0, 13, 0, 14, "REFERENCIA", flexAlignCenterCenter
            FORMATO_CELDA Fg1, 0, 13, vbBlack, True, &HD8E9EC
            
            .ColWidth(1) = 350
    '
            .TextMatrix(1, 1) = "N° Registro":  .ColWidth(1) = 900:   .ColAlignment(1) = flexAlignLeftCenter:     .Row = 1: .Col = 1: .CellAlignment = flexAlignLeftCenter
            .TextMatrix(1, 2) = "Origen":       .ColWidth(2) = 1200:   .ColAlignment(2) = flexAlignLeftCenter:     .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftCenter
            .TextMatrix(1, 3) = "T.D.":         .ColWidth(3) = 450:    .ColAlignment(3) = flexAlignLeftCenter:     .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
            .TextMatrix(1, 4) = "N°.Documento": .ColWidth(4) = 1600:   .ColAlignment(4) = flexAlignLeftCenter:     .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftCenter
            .TextMatrix(1, 5) = "Fch.Emi.":     .ColWidth(5) = 800:    .ColAlignment(5) = flexAlignCenterBottom:   .Row = 1: .Col = 5: .CellAlignment = flexAlignCenterBottom
            .TextMatrix(1, 6) = "Fch.Ven.":     .ColWidth(6) = 800:    .ColAlignment(6) = flexAlignCenterBottom:   .Row = 1: .Col = 6: .CellAlignment = flexAlignCenterBottom
            .TextMatrix(1, 7) = "M":            .ColWidth(7) = 450:    .ColAlignment(7) = flexAlignLeftCenter:    .Row = 1: .Col = 7: .CellAlignment = flexAlignCenterBottom
            
            .TextMatrix(1, 8) = "Imp":          .ColWidth(8) = 900:    .ColAlignment(8) = flexAlignRightCenter:   .Row = 1: .Col = 8: .CellAlignment = flexAlignRightCenter
            .TextMatrix(1, 9) = "T.C.":         .ColWidth(9) = 500:    .ColAlignment(9) = flexAlignRightCenter:    .Row = 1: .Col = 9: .CellAlignment = flexAlignRightCenter
            '----------------------
            .TextMatrix(1, 10) = "Debe":       .ColWidth(10) = 1150:  .ColAlignment(10) = flexAlignRightCenter:   .Row = 1: .Col = 10: .CellAlignment = flexAlignRightCenter
            .TextMatrix(1, 11) = "Haber":       .ColWidth(11) = 1150:  .ColAlignment(11) = flexAlignRightCenter:   .Row = 1: .Col = 11: .CellAlignment = flexAlignRightCenter
            .TextMatrix(1, 12) = "Saldo":       .ColWidth(12) = 1150:  .ColAlignment(12) = flexAlignRightCenter:   .Row = 1: .Col = 12: .CellAlignment = flexAlignRightCenter
            
            .TextMatrix(1, 13) = "N°.Documento":       .ColWidth(13) = 1400:  .ColAlignment(13) = flexAlignLeftCenter:   .Row = 1: .Col = 13: .CellAlignment = flexAlignLeftCenter
            .TextMatrix(1, 14) = "Glosa":       .ColWidth(14) = 3500:  .ColAlignment(14) = flexAlignLeftCenter:   .Row = 1: .Col = 14: .CellAlignment = flexAlignLeftCenter
            
            
            .TextMatrix(1, 15) = "N°. Cuenta":       .ColWidth(15) = 0:  .ColAlignment(15) = flexAlignLeftCenter:   .Row = 1: .Col = 15: .CellAlignment = flexAlignLeftCenter
            .TextMatrix(1, 16) = "Descripción":       .ColWidth(16) = 0:  .ColAlignment(16) = flexAlignLeftCenter:   .Row = 1: .Col = 16: .CellAlignment = flexAlignLeftCenter
            
            .FrozenCols = 0
            .SelectionMode = flexSelectionByRow
        End With
        
        With Fg2
            '-----
            .Rows = 1
            .Cols = 6
            .FixedRows = 1
            .FrozenCols = 0
            .RowHeight(0) = 250
            .ColWidth(0) = 200:
            .SelectionMode = flexSelectionByRow
            .ExplorerBar = flexExSortShow
            .TextMatrix(0, 1) = "R.U.C.":   .ColWidth(1) = 1200:  .ColAlignment(1) = flexAlignCenterCenter: .Row = 0: .Col = 1: .CellAlignment = flexAlignCenterCenter
            .TextMatrix(0, 2) = "Nombres":  .ColWidth(2) = 5500:  .ColAlignment(2) = flexAlignLeftCenter:   .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftCenter
            '----------------------
            .TextMatrix(0, 3) = "Debe":    .ColWidth(3) = 1300:  .ColAlignment(3) = flexAlignRightCenter:  .Row = 0: .Col = 3: .CellAlignment = flexAlignRightCenter
            .TextMatrix(0, 4) = "Haber":    .ColWidth(4) = 1300:  .ColAlignment(4) = flexAlignRightCenter:  .Row = 0: .Col = 4: .CellAlignment = flexAlignRightCenter
            
            .TextMatrix(0, 5) = "Saldo":    .ColWidth(5) = 1300:  .ColAlignment(5) = flexAlignRightCenter:  .Row = 0: .Col = 5: .CellAlignment = flexAlignRightCenter
            For A = 1 To .Cols - 1
                FORMATO_CELDA Fg2, 0, A, vbBlack, True, &HD8E9EC
            Next
            .SelectionMode = flexSelectionByRow
        End With
        
        TabOne1.TabVisible(1) = True
        
    Else
        With Fg1
            '-----
            .Rows = 2
            .Cols = 24
            .FixedRows = 2
            .FixedCols = 1
            .FrozenCols = 0
            .RowHeight(0) = 250
            .RowHeight(1) = 250
            .ColWidth(0) = 200
            .SelectionMode = flexSelectionByRow
            .ExplorerBar = flexExSortShow
            UNIR_CELDAS Fg1, 0, 1, 0, 10, "DATOS DEL DOCUMENTO", flexAlignCenterCenter
            FORMATO_CELDA Fg1, 0, 1, vbBlack, True, &HD8E9EC
            
            
            UNIR_CELDAS Fg1, 0, 11, 0, 18, "REFERENCIA", flexAlignCenterCenter
            FORMATO_CELDA Fg1, 0, 11, vbBlack, True, &HD8E9EC
            
            .ColWidth(1) = 350
    
            .TextMatrix(1, 1) = "N° RUC":       .ColWidth(1) = 900:    .ColAlignment(1) = flexAlignLeftCenter:     .Row = 1: .Col = 1: .CellAlignment = flexAlignLeftCenter
            .TextMatrix(1, 2) = "Nombres":      .ColWidth(2) = 1500:   .ColAlignment(2) = flexAlignLeftCenter:     .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftCenter
            .TextMatrix(1, 3) = "N° Registro":  .ColWidth(3) = 900:    .ColAlignment(3) = flexAlignLeftCenter:     .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
            .TextMatrix(1, 4) = "T.D.":         .ColWidth(4) = 450:    .ColAlignment(4) = flexAlignLeftCenter:     .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftCenter
            .TextMatrix(1, 5) = "N°.Documento": .ColWidth(5) = 1500:   .ColAlignment(5) = flexAlignLeftCenter:     .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftCenter
            .TextMatrix(1, 6) = "Fch.Emi.":     .ColWidth(6) = 800:    .ColAlignment(6) = flexAlignCenterCenter:   .Row = 1: .Col = 6: .CellAlignment = flexAlignCenterCenter
            .TextMatrix(1, 7) = "M":            .ColWidth(7) = 450:    .ColAlignment(7) = flexAlignLeftCenter:     .Row = 1: .Col = 7: .CellAlignment = flexAlignCenterCenter
            .TextMatrix(1, 8) = "T.C.":         .ColWidth(8) = 500:    .ColAlignment(8) = flexAlignRightCenter:    .Row = 1: .Col = 8: .CellAlignment = flexAlignRightCenter
            .TextMatrix(1, 9) = "Imp":          .ColWidth(9) = 900:    .ColAlignment(9) = flexAlignRightCenter:    .Row = 1: .Col = 9: .CellAlignment = flexAlignRightCenter
            .TextMatrix(1, 10) = "Glosa":       .ColWidth(10) = 1000:  .ColAlignment(10) = flexAlignLeftCenter:    .Row = 1: .Col = 10: .CellAlignment = flexAlignLeftCenter
            '----------------------
            .TextMatrix(1, 11) = "Tot.Reg":      .ColWidth(11) = 700:   .ColAlignment(11) = flexAlignCenterCenter:    .Row = 1: .Col = 11: .CellAlignment = flexAlignCenterCenter
            .TextMatrix(1, 12) = "Fch.Cancel":   .ColWidth(12) = 900:   .ColAlignment(12) = flexAlignCenterCenter:  .Row = 1: .Col = 12: .CellAlignment = flexAlignCenterCenter
            .TextMatrix(1, 13) = "N° Registro":  .ColWidth(13) = 900:   .ColAlignment(13) = flexAlignRightCenter:   .Row = 1: .Col = 13: .CellAlignment = flexAlignLeftCenter
            .TextMatrix(1, 14) = "Nombres":      .ColWidth(14) = 1200:  .ColAlignment(14) = flexAlignLeftCenter:    .Row = 1: .Col = 14: .CellAlignment = flexAlignLeftCenter
            .TextMatrix(1, 15) = "T.D.":         .ColWidth(15) = 450:   .ColAlignment(15) = flexAlignLeftCenter:    .Row = 1: .Col = 15: .CellAlignment = flexAlignLeftCenter
            .TextMatrix(1, 16) = "N°.Documento": .ColWidth(16) = 1200:  .ColAlignment(16) = flexAlignLeftCenter:    .Row = 1: .Col = 16: .CellAlignment = flexAlignLeftCenter
            .TextMatrix(1, 17) = "M":            .ColWidth(17) = 550:   .ColAlignment(17) = flexAlignRightCenter:   .Row = 1: .Col = 17: .CellAlignment = flexAlignRightCenter
            .TextMatrix(1, 18) = "Total":        .ColWidth(18) = 1150:  .ColAlignment(18) = flexAlignRightCenter:   .Row = 1: .Col = 18: .CellAlignment = flexAlignRightCenter
            .TextMatrix(1, 19) = "Glosa":        .ColWidth(19) = 2500:  .ColAlignment(19) = flexAlignLeftCenter:    .Row = 1: .Col = 19: .CellAlignment = flexAlignLeftCenter

            .TextMatrix(1, 20) = "Saldo":        .ColWidth(20) = 1150:  .ColAlignment(20) = flexAlignRightCenter:   .Row = 1: .Col = 20: .CellAlignment = flexAlignRightCenter
            .TextMatrix(1, 21) = "IdDoc":        .ColWidth(21) = 0:     .ColAlignment(21) = flexAlignLeftCenter:    .Row = 1: .Col = 21: .CellAlignment = flexAlignLeftCenter
            .TextMatrix(1, 22) = "IdMon":        .ColWidth(22) = 0:     .ColAlignment(22) = flexAlignLeftCenter:    .Row = 1: .Col = 22: .CellAlignment = flexAlignLeftCenter
            
            .TextMatrix(1, 23) = "Doc. Ref.":    .ColWidth(23) = 1500:  .ColAlignment(23) = flexAlignLeftCenter:    .Row = 1: .Col = 23: .CellAlignment = flexAlignLeftCenter
            
            .FrozenCols = 10
            
            
        End With
        
        With Fg2
            '-----
            .Rows = 2
            .Cols = 9
            .FixedRows = 2
            .FrozenCols = 0
            .RowHeight(0) = 250
            .ColWidth(0) = 200:
            .ExplorerBar = flexExSortShow
            
            UNIR_CELDAS Fg2, 0, 1, 0, 4, "DATOS DEL DOCUMENTO", flexAlignCenterCenter
            FORMATO_CELDA Fg2, 0, 1, vbBlack, True, &HD8E9EC
            
            UNIR_CELDAS Fg2, 0, 5, 0, 6, "COBRANZA / PAGO ", flexAlignCenterCenter
            FORMATO_CELDA Fg2, 0, 5, vbBlack, True, &HD8E9EC
            
            UNIR_CELDAS Fg2, 0, 7, 0, 8, "SALDOS", flexAlignCenterCenter
            FORMATO_CELDA Fg2, 0, 7, vbBlack, True, &HD8E9EC
            
            .TextMatrix(1, 1) = "R.U.C.":   .ColWidth(1) = 1200:  .ColAlignment(1) = flexAlignCenterCenter: .Row = 0: .Col = 1: .CellAlignment = flexAlignCenterCenter
            .TextMatrix(1, 2) = "Nombres":  .ColWidth(2) = 4500:  .ColAlignment(2) = flexAlignLeftCenter:   .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftCenter
            '----------------------
            .TextMatrix(1, 3) = "Imp MN":    .ColWidth(3) = 1300:  .ColAlignment(3) = flexAlignRightCenter:  .Row = 0: .Col = 3: .CellAlignment = flexAlignRightCenter
            .TextMatrix(1, 4) = "Imp ME":    .ColWidth(4) = 1300:  .ColAlignment(4) = flexAlignRightCenter:  .Row = 0: .Col = 4: .CellAlignment = flexAlignRightCenter
            
            .TextMatrix(1, 5) = "Imp MN":    .ColWidth(5) = 1300:  .ColAlignment(5) = flexAlignRightCenter:  .Row = 0: .Col = 5: .CellAlignment = flexAlignRightCenter
            .TextMatrix(1, 6) = "Imp ME":    .ColWidth(6) = 1300:  .ColAlignment(6) = flexAlignRightCenter:  .Row = 0: .Col = 6: .CellAlignment = flexAlignRightCenter
            
            .TextMatrix(1, 7) = "Imp MN":    .ColWidth(7) = 1300:  .ColAlignment(7) = flexAlignRightCenter:  .Row = 0: .Col = 7: .CellAlignment = flexAlignRightCenter
            .TextMatrix(1, 8) = "Imp ME":    .ColWidth(8) = 1300:  .ColAlignment(8) = flexAlignRightCenter:  .Row = 0: .Col = 8: .CellAlignment = flexAlignRightCenter
            
            For A = 1 To .Cols - 1
                FORMATO_CELDA Fg2, 0, A, vbBlack, True, &HD8E9EC
            Next
            .SelectionMode = flexSelectionByRow
        End With
        
'        TabOne1.TabVisible(1) = False
    End If
    
    FraDetalle.Visible = False
    
    Fg4.ColWidth(0) = 0
    Fg4.ColWidth(2) = 0
    Fg4.RowHeight(0) = 0
    
    DoEvents
End Sub

Private Sub pImprimir()
    On Error GoTo error
    Dim oPrint As New SGI2_funciones.formularios
    Dim nPeriodo As String
    Dim nTitulo As String
    Dim nTitulo1 As String
    Dim nTipo As String
    nPeriodo = "Al  " + CStr(TxtFchIni.Valor)
    nTitulo1 = "(Expresado en " & LblMoneda.Caption & ")"
    nTipo = Fg4.TextMatrix(Fg4.Row, 2)
    Me.MousePointer = vbHourglass
    
    If TabOne1.CurrTab = 0 Then
        nTitulo = "Detalle de Cuenta Corriente - " + nTipo
        oPrint.Imprimir_x_VSFlexGrid Fg1, nTitulo, nTitulo1, nPeriodo, True, True
    Else
        nTitulo = "Resumen de Cuenta Corriente - " + nTipo
        oPrint.Imprimir_x_VSFlexGrid Fg2, nTitulo, nTitulo1, nPeriodo, True, True
    End If
    
    Set oPrint = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"
End Sub

Private Sub pConsultar()

    If OptSel1.Value = True Then
        If Chk.Value = 0 Then
            CargarCli4 0
        Else
            VerLineal
        End If
    End If
    
    If OptSel2.Value = True Then
        TabOne2.CurrTab = 0
        If NulosC(TxtCliPro.Text) = "" Then
        
            Select Case NulosN(Fg4.TextMatrix(Fg4.Row, 2))
                Case 1, 40, 999 '--Compras
                    MsgBox "No ha especificado el proveedor a consultar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Case 9  '--Boleta Pago
                    MsgBox "No ha especificado el personal a consultar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Case 2, 37, 41 '--Letras
                    MsgBox "No ha especificado el cliente a consultar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Case 42  '--Planilla letras
                    MsgBox "No ha especificado la entidad bancaria a consultar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                
            End Select
            
            TxtCliPro.SetFocus
            Exit Sub
        End If
        
        If Chk.Value = 0 Then
            CargarCli4 NulosN(LblIdCliPro.Caption)
        Else
            VerLineal
        End If
        
    End If

End Sub

Private Sub VerLineal()
    Dim xRstLineal As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLDiario As String
    Dim xBase As String
    Dim nSQLPer As String
    Dim nSQLFecha As String
    Dim nSQLDiario1 As String
    Dim xImpSaldo As Double
    Dim nSQLWhere As String
    Dim nSQLApertura As String '--filtro para documentos de apertura
    Dim xEliminaReg As Boolean
    Dim nSQLDocNCBancos As String '--almacenara los documentos de NC que pasan x banco
   '--2 Ventas
   '--1 Compras
   '--9 Boleta
   
   '-----------------------------------------------
    If Fg4.Row < 1 Then Exit Sub
    xIdLibro = Fg4.TextMatrix(Fg4.Row, 2)
    Select Case xIdLibro
        Case 1, 4, 40, 9
            xIdLibroRef = "6,8,39"
        Case 2 '--ventas
            xIdLibroRef = "5,6,8,37"
        Case 37 '--Letras
            xIdLibroRef = "6,42"
        Case 41 '--Lgd
            xIdLibroRef = "6,41"
        Case 42 '--Planilla Letras
            xIdLibroRef = "6"
        Case 999
            If OptReem1.Value = True Then
                xIdLibroRef = "6"
            Else
                xIdLibroRef = "41"
            End If
    End Select
    '-----------------------------------------------
    TabOne1.CurrTab = 1
    Fg1.Rows = Fg1.FixedRows
    Fg2.Rows = Fg2.FixedRows
    DoEvents
   
    '--------------------------------------------
    '--con_diario.tipmov(1=Ingresos; 2=Egresos)
    '--con_diario.tipo  (1=Origen;   2=Destino)
    '--con_diario.rtipdoc=7(Nota de Credito)
    If xIdLibro = 1 Or xIdLibro = 4 Or xIdLibro = 9 Or xIdLibro = 40 Or xIdLibro = 999 Then
    '--compras, honorarios, reembolsables, boleta pago,percepciones
        xBase = vbCr & " IIf(con_diario.tipmov =1, IIf(con_diario.tipo =1,IIf(con_diario.rtipdoc=7,-1,1),IIf(con_diario.rtipdoc=7,1,-1)) , IIf(con_diario.tipo in (1),IIf(con_diario.rtipdoc=7,1,-1),1) ) as xbase, "
    Else
    '--ventas, lgd, letras, planilla letras
        xBase = " IIf(con_diario.tipmov =1, IIf(con_diario.tipo =1,IIf(con_diario.rtipdoc=7,1,-1),1), IIf(con_diario.tipo in (0,1),1,IIf(con_diario.rtipdoc=7,1,-1)) ) as xbase, "
    End If
    '--------------------------------------------
    If OptFch(0).Value = True Then '--x fecha de documento
        nSQLFecha = " and ( vta_ventas.fchdoc between CDate('" & TxtFchIni.Valor & "') and CDate('" & TxtFchFin.Valor & "') )"
    ElseIf OptFch(1).Value = True Then '--x fecha de registro
        nSQLFecha = " and ( vta_ventas.fchreg between CDate('" & TxtFchIni.Valor & "') and CDate('" & TxtFchFin.Valor & "') )"
    End If
    
    If NulosN(LblIdCliPro.Caption) <> 0 Then
        nSQLPer = " and con_diario.ridper =" & NulosN(LblIdCliPro.Caption)
        If OptAperturaCon.Value = True Then nSQLApertura = " or (vta_ventas.numreg='000001' " & nSQLPer & " ) "
    Else
        If OptAperturaCon.Value = True Then nSQLApertura = " or vta_ventas.numreg='000001' "
    End If
    If OptAperturaSin.Value = True Then nSQLApertura = " and vta_ventas.numreg<>'000001' "
    If OptAperturaSolo.Value = True Then nSQLApertura = " and vta_ventas.numreg='000001' "
    '--------------------------------------------
    nSQLWhere = nSQLFecha & nSQLPer & nSQLApertura
    '--------------------------------------------
    
    
    If xIdLibro = 1 Then
        
        '--Verificar si hay documentos de NC que fueron registrados en Tesoreria Ingresos - Egresos
        nSQLDocNCBancos = BuscarNCBancos()
        If nSQLDocNCBancos <> "" Then
            nSQLDocNCBancos = " and com_compras.id not in (" & nSQLDocNCBancos & ")"
        End If
        
        '--Cancelacion de compras con nota de credito excepto nc que se registran en tesoreria
        nSQLDiario1 = " UNION " _
            + vbCr + " SELECT Left(com_compras_1.numreg,2) & mae_libros_1.codsun & Right(com_compras_1.numreg,4) AS rregistro, Mid(com_compras!numreg,1,2) & mae_libros!codsun & Mid(com_compras.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_prov.nombre AS razonsocial, mae_documento.abrev, " _
            + vbCr + " com_compras.numser & '-' & com_compras.numdoc AS numdoc, com_compras.fchdoc AS fchemi, mae_moneda.simbolo, " _
            + vbCr + " IIf(com_compras_1.tc=0,con_tc.impven,com_compras_1.tc) AS tipcam, 2 AS tipmov, 1 AS tipo, 1 AS xbase, com_compras.imptot AS imptotal, " _
            + vbCr + " IIf(com_compras.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, " _
            + vbCr + " IIf(com_compras.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) AS imptotdol, " _
            + vbCr + " com_compras.idpro AS ridper, com_compras_1.numser & '-' & com_compras_1.numdoc AS numdoc2, com_compras.glosa AS glosaope, com_compras_1.id AS iddoc, com_compras.idmon " _
            + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((((com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) INNER JOIN com_compras AS com_compras_1 ON com_compras.iddocref = com_compras_1.id)  " _
            + vbCr + " LEFT JOIN mae_libros AS mae_libros_1 ON com_compras_1.idlib = mae_libros_1.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) LEFT JOIN mae_prov ON com_compras.idpro = mae_prov.id " _
            + vbCr + " WHERE (com_compras.iddocref<>0 ) " & Replace(nSQLPer, "con_diario.ridper", "com_compras.idpro") & nSQLDocNCBancos

        '--Cancelacion de notas de credito con compras excepto nc que se registran en tesoreria
        nSQLDocNCBancos = Replace(nSQLDocNCBancos, "com_compras", "com_compras_1")

        nSQLDiario1 = nSQLDiario1 _
            + vbCr + " UNION " _
            + vbCr + " SELECT Left(com_compras_1.numreg,2) & mae_libros_1.codsun & Right(com_compras_1.numreg,4) AS rregistro, Mid(com_compras!numreg,1,2) & mae_libros!codsun & Mid(com_compras.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_prov.nombre AS razonsocial, mae_documento.abrev, " _
            + vbCr + " com_compras.numser & '-' & com_compras.numdoc AS numdoc, com_compras.fchdoc AS fchemi, mae_moneda.simbolo, " _
            + vbCr + " IIf(com_compras_1.tc=0,con_tc.impven,com_compras_1.tc) AS tipcam, 2 AS tipmov, 1 AS tipo, 1 AS xbase, com_compras_1.imptot AS imptotal, " _
            + vbCr + " IIf(com_compras.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, " _
            + vbCr + " IIf(com_compras.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) AS imptotdol, " _
            + vbCr + " com_compras.idpro AS ridper, com_compras_1.numser & '-' & com_compras_1.numdoc AS numdoc2, com_compras.glosa AS glosaope, com_compras_1.id AS iddoc, com_compras.idmon " _
            + vbCr + " FROM (com_compras AS com_compras_1 LEFT JOIN mae_libros AS mae_libros_1 ON com_compras_1.idlib = mae_libros_1.id) INNER JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (((com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) " _
            + vbCr + " LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) LEFT JOIN mae_prov ON com_compras.idpro = mae_prov.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON com_compras_1.iddocref = com_compras.id " _
            + vbCr + " WHERE (com_compras_1.iddocref<>0 ) " & Replace(nSQLPer, "con_diario.ridper", "com_compras.idpro") & nSQLDocNCBancos
        '--------------------------------------------------
    
    ElseIf xIdLibro = 2 Then
        '--Verificar si hay documentos de NC que fueron registrados en Tesoreria Ingresos - Egresos
        nSQLDocNCBancos = BuscarNCBancos()
        If nSQLDocNCBancos <> "" Then
            nSQLDocNCBancos = " and vta_ventas.id not in (" & nSQLDocNCBancos & ")"
        End If
        
        '--Cancelacion de ventas con nota de credito excepto nc que se registran en tesoreria
        nSQL = nSQL _
            + vbCr + " UNION " _
            + vbCr + " SELECT Left(vta_ventas_1.numreg,2) & mae_libros_1.codsun & Right(vta_ventas_1.numreg,4) AS rregistro, Mid(vta_ventas!numreg,1,2) & mae_libros!codsun & Mid(vta_ventas.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_prov.nombre AS razonsocial, mae_documento.abrev, " _
            + vbCr + " vta_ventas.numser & '-' & vta_ventas.numdoc AS numdoc, vta_ventas.fchdoc AS fchemi, mae_moneda.simbolo, IIf(vta_ventas_1.tc=0,con_tc.impven,vta_ventas_1.tc) AS tipcam, " _
            + vbCr + " 2 AS tipmov, 1 AS tipo, 1 AS xbase, vta_ventas.imptotdoc AS imptotal, " _
            + vbCr + " IIf(vta_ventas.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, " _
            + vbCr + " IIf(vta_ventas.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) AS imptotdol, " _
            + vbCr + " vta_ventas.idcli AS ridper, vta_ventas_1.numser & '-' & vta_ventas_1.numdoc AS numdoc2, vta_ventas.glosa AS glosaope, vta_ventas_1.id AS iddoc, vta_ventas.idmon " _
            + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((((vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) INNER JOIN vta_ventas AS vta_ventas_1 ON vta_ventas.iddocref = vta_ventas_1.id)  " _
            + vbCr + " LEFT JOIN mae_libros AS mae_libros_1 ON vta_ventas_1.idlib = mae_libros_1.id) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) LEFT JOIN mae_prov ON vta_ventas.idcli = mae_prov.id " _
            + vbCr + " WHERE (vta_ventas.iddocref<>0 ) " & Replace(nSQLPer, "con_diario.ridper", "vta_ventas.idcli") & nSQLDocNCBancos

        '--Cancelacion de notas de credito con ventas excepto nc que se registran en tesoreria
        nSQLDocNCBancos = Replace(nSQLDocNCBancos, "vta_ventas", "vta_ventas_1")

        nSQL = nSQL _
            + vbCr + " UNION " _
            + vbCr + " SELECT Left(vta_ventas_1.numreg,2) & mae_libros_1.codsun & Right(vta_ventas_1.numreg,4) AS rregistro, Mid(vta_ventas!numreg,1,2) & mae_libros!codsun & Mid(vta_ventas.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_prov.nombre AS razonsocial, mae_documento.abrev, " _
            + vbCr + " vta_ventas.numser & '-' & vta_ventas.numdoc AS numdoc, vta_ventas.fchdoc AS fchemi, mae_moneda.simbolo, IIf(vta_ventas_1.tc=0,con_tc.impven,vta_ventas_1.tc) AS tipcam,  " _
            + vbCr + " 2 AS tipmov, 1 AS tipo, 1 AS xbase, vta_ventas_1.imptotdoc AS imptotal, " _
            + vbCr + " IIf(vta_ventas.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, " _
            + vbCr + " IIf(vta_ventas.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) AS imptotdol, " _
            + vbCr + " vta_ventas.idcli AS ridper, vta_ventas_1.numser & '-' & vta_ventas_1.numdoc AS numdoc2, vta_ventas.glosa AS glosaope, vta_ventas_1.id AS iddoc, vta_ventas.idmon " _
            + vbCr + " FROM (vta_ventas AS vta_ventas_1 LEFT JOIN mae_libros AS mae_libros_1 ON vta_ventas_1.idlib = mae_libros_1.id) INNER JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (((vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) " _
            + vbCr + " LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_prov ON vta_ventas.idcli = mae_prov.id) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON vta_ventas_1.iddocref = vta_ventas.id " _
            + vbCr + " WHERE (vta_ventas_1.iddocref<>0 ) " & Replace(nSQLPer, "con_diario.ridper", "vta_ventas.idcli") & nSQLDocNCBancos
        '--------------------------------------------------
    
    End If
    
    nSQLDiario = " SELECT Last(xdet.rregistro) as xrregistro, xdet.iddoc, last(xdet.simbolo) as xsimbolo, Sum(xdet.imptotsol) AS xtotsol, Sum(xdet.imptotdol) AS xtotdol, Last(xdet.registro) AS xregistro, Last(xdet.fchemi) AS xfchcancel, Count(xdet.tipmov) AS xcanreg,last(xdet.rglosaope) as xglosa ,last(xdet.razonsocial) as xnombre, last(xdet.numdoc) as xnumdoc,last(xdet.abrev) as xabrev  " _
        + vbCr + " FROM ( " _
        + vbCr + " SELECT con_diario.rregistro, Format([con_diario].[idmes],'00') & [mae_libros].[codsun] & Format([con_diario].[numasi],'0000') AS registro, mae_libros.descripcion AS libro, IIf([con_diario].[ridtipper2]=5,[mae_bancos].[abrev],IIf([con_diario].[ridtipper2]=2,[mae_cliente].[nombre],IIf([con_diario].[ridtipper2]=1,[mae_prov].[nombre],''))) AS razonsocial, tes_documentos.abrev, con_diario.rnumerodoc2 AS numdoc, con_diario.rfchope2 AS fchemi, mae_moneda.simbolo, IIf([con_diario].[aplicatc]=0,[con_tc].[impven],[con_diario].[tc]) AS tipcam, " _
        + vbCr + " con_diario.tipmov, con_diario.tipo, " & xBase _
        + vbCr + " IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS imptotal, " _
        + vbCr + " IIf(con_diario.idmon=1,imptotal,imptotal*tipcam) * xbase  AS imptotsol, " _
        + vbCr + " IIf(con_diario.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) * xbase AS imptotdol, " _
        + vbCr + " con_diario.ridper, con_diario.rnumerodoc AS numdoc2, con_diario.rglosaope, con_diario.iddoc, con_diario.idmon  " _
        + vbCr + " FROM ((((((con_diario LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN tes_documentos ON con_diario.rtipdoc2 = tes_documentos.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN mae_bancos ON con_diario.ridper2 = mae_bancos.id) LEFT JOIN mae_cliente ON con_diario.ridper2 = mae_cliente.id) LEFT JOIN mae_prov ON con_diario.ridper2 = mae_prov.id " _
        + vbCr + " WHERE (((con_diario.idlib) In (" & xIdLibroRef & ")) AND ((con_diario.ridlib)=" & xIdLibro & ")) " & nSQLPer _
        + vbCr + nSQLDiario1 _
        + vbCr + " ) AS xdet " _
        + vbCr + " GROUP BY xdet.iddoc "

    Select Case xIdLibro
        Case 1 '--compras
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "com_compras")
            nSQLWhere = Replace(nSQLWhere, "con_diario.ridper", "com_compras.idpro")
                        
            nSQL = "SELECT  com_compras.id as iddoc,com_compras.tipdoc, mae_prov.numruc, mae_prov.nombre AS nombre, IIf(com_compras.numreg Is Null Or com_compras.numreg='',mae_libros.codsun,Left(com_compras.numreg,2) & mae_libros.codsun & Right(com_compras.numreg,4)) AS registro, mae_documento.abrev, " _
                + vbCr + " com_compras.numser+'-'+com_compras.numdoc AS numdoc2, com_compras.fchdoc, mae_moneda.simbolo, " _
                + vbCr + " IIf(com_compras.tc Is Null Or com_compras.tc=0,con_tc.impven,com_compras.tc) AS tipcam, com_compras.idmon, IIf(com_compras.numreg='000001',com_compras.imptotori,com_compras.imptot) AS imptotal, com_compras.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(com_compras.idmon=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(com_compras.idmon=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, " _
                + vbCr + " com_compras.glosa , com_compras.numerodocref as docref, " _
                + vbCr + " xtotsol, xtotdol, IIf(com_compras.idmon=1,imptotal-iif(xtotsol is null,0,xtotsol),imptotal-iif(xtotdol is null,0,xtotdol)) AS saldo, xregistro, xnombre,xabrev,xnumdoc,xfchcancel,xglosa ,xcanreg , xsimbolo " _
                + vbCr + " FROM ( (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN (mae_documento RIGHT JOIN (com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_documento.id = com_compras.tipdoc) ON mae_prov.id = com_compras.idpro) ON mae_moneda.id = com_compras.idmon) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
                + vbCr + " ) LEFT JOIN ( " & nSQLDiario & "" _
                + vbCr + " ) AS xpag ON com_compras.id = xpag.iddoc " _
                + vbCr + " WHERE (((IIf(com_compras.numreg='000001',com_compras.imptotori,com_compras.imptot))<>0)) " & nSQLWhere _
                + vbCr + " ORDER BY IIf(com_compras.numreg Is Null Or com_compras.numreg='',mae_libros.codsun,Left(com_compras.numreg,2) & mae_libros.codsun & Right(com_compras.numreg,4)), com_compras.fchdoc;"
                
        Case 999 '--Reembolsables
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "com_reembolsables")
            nSQLWhere = Replace(nSQLWhere, "con_diario.ridper", "com_reembolsables.idpro")
            nSQLWhere = Replace(nSQLWhere, "fchreg", "fchdoc")
            
            nSQL = "SELECT  com_reembolsables.id as iddoc,com_reembolsables.tipdoc, mae_prov.numruc, mae_prov.nombre AS nombre, IIf(com_reembolsables.numreg Is Null Or com_reembolsables.numreg='',mae_libros.codsun,Left(com_reembolsables.numreg,2) & mae_libros.codsun & Right(com_reembolsables.numreg,4)) AS registro, mae_documento.abrev, " _
                + vbCr + " com_reembolsables.numser+'-'+com_reembolsables.numdoc AS numdoc2, com_reembolsables.fchdoc, mae_moneda.simbolo, " _
                + vbCr + " IIf(com_reembolsables.tc Is Null Or com_reembolsables.tc=0,con_tc.impven,com_reembolsables.tc) AS tipcam, com_reembolsables.idmon, IIf(com_reembolsables.numreg='000001',com_reembolsables.imptotori,com_reembolsables.imptot) AS imptotal, com_reembolsables.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(com_reembolsables.idmon=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(com_reembolsables.idmon=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, " _
                + vbCr + " com_reembolsables.glosa , com_reembolsables.numerodocref as docref, " _
                + vbCr + " xtotsol, xtotdol, IIf(com_reembolsables.idmon=1,imptotal-xtotsol,imptotal-xtotdol) AS saldo, xregistro, xnombre,xabrev,xnumdoc,xfchcancel,xglosa ,xcanreg , xsimbolo " _
                + vbCr + " FROM ( (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN (mae_documento RIGHT JOIN (com_reembolsables LEFT JOIN mae_libros ON com_reembolsables.idlib = mae_libros.id) ON mae_documento.id = com_reembolsables.tipdoc) ON mae_prov.id = com_reembolsables.idpro) ON mae_moneda.id = com_reembolsables.idmon) LEFT JOIN con_tc ON com_reembolsables.fchdoc = con_tc.fecha " _
                + vbCr + " ) LEFT JOIN ( " & nSQLDiario & "" _
                + vbCr + " ) AS xpag ON com_reembolsables.id = xpag.iddoc " _
                + vbCr + " WHERE (((com_reembolsables.tipdoc)<>7) AND ((IIf(com_reembolsables.numreg='000001',com_reembolsables.imptotori,com_reembolsables.imptot))<>0)) " & nSQLWhere _
                + vbCr + " ORDER BY IIf(com_reembolsables.numreg Is Null Or com_reembolsables.numreg='',mae_libros.codsun,Left(com_reembolsables.numreg,2) & mae_libros.codsun & Right(com_reembolsables.numreg,4)), com_reembolsables.fchdoc;"
        
        Case 2 '--ventas
            nSQLWhere = Replace(nSQLWhere, "con_diario.ridper", "vta_ventas.idcli")
            
            nSQL = "SELECT  vta_ventas.id as iddoc,vta_ventas.tipdoc, mae_cliente.numruc, mae_cliente.nombre AS nombre, IIf(vta_ventas.numreg Is Null Or vta_ventas.numreg='',mae_libros.codsun,Left(vta_ventas.numreg,2) & mae_libros.codsun & Right(vta_ventas.numreg,4)) AS registro, mae_documento.abrev, " _
                + vbCr + " vta_ventas.numser+'-'+vta_ventas.numdoc AS numdoc2, vta_ventas.fchdoc, mae_moneda.simbolo, " _
                + vbCr + " IIf(vta_ventas.tc Is Null Or vta_ventas.tc=0,con_tc.impven,vta_ventas.tc) AS tipcam, vta_ventas.idmon, IIf(vta_ventas.numreg='000001',vta_ventas.imptotori,vta_ventas.imptotdoc) AS imptotal, vta_ventas.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(vta_ventas.idmon=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(vta_ventas.idmon=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, " _
                + vbCr + " vta_ventas.glosa ,vta_ventas.numerodocref as docref,  " _
                + vbCr + " xtotsol, xtotdol, IIf(vta_ventas.idmon=1,imptotal-xtotsol,imptotal-xtotdol) AS saldo,xregistro, xnombre,xabrev,xnumdoc,xfchcancel,xglosa ,xcanreg , xsimbolo " _
                + vbCr + " FROM ( (mae_moneda RIGHT JOIN (mae_cliente RIGHT JOIN (mae_documento RIGHT JOIN (vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) ON mae_documento.id = vta_ventas.tipdoc) ON mae_cliente.id = vta_ventas.idcli) ON mae_moneda.id = vta_ventas.idmon) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha " _
                + vbCr + " ) LEFT JOIN ( " & nSQLDiario & "" _
                + vbCr + " ) AS xpag ON vta_ventas.id = xpag.iddoc " _
                + vbCr + " WHERE vta_ventas.anulado=0 and (IIf(vta_ventas.numreg='000001',vta_ventas.imptotori,vta_ventas.imptotdoc)<>0) " & nSQLWhere _
                + vbCr + " ORDER BY IIf(vta_ventas.numreg Is Null Or vta_ventas.numreg='',mae_libros.codsun,Left(vta_ventas.numreg,2) & mae_libros.codsun & Right(vta_ventas.numreg,4)), vta_ventas.fchdoc;"
        
        Case 4 '--Percepciones
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "con_percepcion")
            nSQLWhere = Replace(nSQLWhere, "con_diario.ridper", "con_percepcion.idcli")
            
            nSQL = "SELECT con_percepcion.id AS iddoc, con_percepcion.tipdoc, mae_prov.numruc, mae_prov.nombre AS nombre, IIf([con_percepcion].[numreg] Is Null Or [con_percepcion].[numreg]='',[mae_libros].[codsun],Left([con_percepcion].[numreg],2) & [mae_libros].[codsun] & Right([con_percepcion].[numreg],4)) AS registro, mae_documento.abrev, " _
                + vbCr + " [con_percepcion].[numser]+'-'+[con_percepcion].[numdoc] AS numdoc2, con_percepcion.fchdoc, mae_moneda.simbolo, " _
                + vbCr + " IIf([con_percepcion].[tc] Is Null Or [con_percepcion].[tc]=0,[con_tc].[impven],[con_percepcion].[tc]) AS tipcam, con_percepcion.idmon, con_percepcion.imptotper AS imptotal, con_percepcion.impsal, " _
                + vbCr + " IIf([imptotal]=0,0,IIf([con_percepcion].[idmon]=1,[imptotal],IIf([tipcam] Is Null,0,[imptotal]*[tipcam]))) AS imptotsol, " _
                + vbCr + " IIf([imptotal]=0,0,IIf([con_percepcion].[idmon]=2,[imptotal],IIf([tipcam] Is Null,0,[imptotal]/[tipcam]))) AS imptotdol, " _
                + vbCr + " con_percepcion.glosa, '' AS docref, " _
                + vbCr + " xtotsol, xtotdol, IIf(con_percepcion.idmon=1,imptotal-xtotsol,imptotal-xtotdol) AS saldo, xregistro, xnombre,xabrev,xnumdoc,xfchcancel,xglosa ,xcanreg , xsimbolo " _
                + vbCr + " FROM ( ((((con_percepcion LEFT JOIN mae_documento ON con_percepcion.tipdoc = mae_documento.id) LEFT JOIN mae_libros ON con_percepcion.idlib = mae_libros.id) LEFT JOIN con_tc ON con_percepcion.fchdoc = con_tc.fecha) LEFT JOIN mae_prov ON con_percepcion.idcli = mae_prov.id) LEFT JOIN mae_moneda ON con_percepcion.idmon = mae_moneda.id " _
                + vbCr + " ) LEFT JOIN ( " & nSQLDiario & "" _
                + vbCr + " ) AS xpag ON con_percepcion.id = xpag.iddoc " _
                + vbCr + " Where (((con_percepcion.imptotper) <> 0)) " & nSQLWhere _
                + vbCr + " ORDER BY IIf([con_percepcion].[numreg] Is Null Or [con_percepcion].[numreg]='',[mae_libros].[codsun],Left([con_percepcion].[numreg],2) & [mae_libros].[codsun] & Right([con_percepcion].[numreg],4)), con_percepcion.fchdoc  "

        Case 9 '--Planilla Pago
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "pla_boleta")
            nSQLWhere = Replace(nSQLWhere, "con_diario.ridper", "pla_boleta.idemp")
                
            nSQL = "SELECT  pla_boleta.id as iddoc,pla_boleta.iddoc as tipdoc, pla_empleados.numdoc as  numruc, pla_empleados.nombre AS nombre, IIf(pla_boleta.numreg Is Null Or pla_boleta.numreg='',mae_libros.codsun,Left(pla_boleta.numreg,2) & mae_libros.codsun & Right(pla_boleta.numreg,4)) AS registro, mae_documento.abrev, " _
                + vbCr + " pla_boleta.numser+'-'+pla_boleta.numdoc AS numdoc2, pla_boleta.fchdoc, mae_moneda.simbolo, " _
                + vbCr + " con_tc.impven AS tipcam, pla_boleta.idmon, pla_boleta.imptot AS imptotal, pla_boleta.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(pla_boleta.idmon=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(pla_boleta.idmon=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, " _
                + vbCr + " pla_boleta.glosa ,'' as docref, " _
                + vbCr + " xtotsol, xtotdol, IIf(pla_boleta.idmon=1,imptotal-iif(xtotsol is null,0,xtotsol),imptotal-iif(xtotdol is null,0,xtotdol)) AS saldo, xregistro, xnombre,xabrev,xnumdoc,xfchcancel,xglosa ,xcanreg , xsimbolo " _
                + vbCr + " FROM ( (mae_moneda RIGHT JOIN (pla_empleados RIGHT JOIN (mae_documento RIGHT JOIN (pla_boleta LEFT JOIN mae_libros ON pla_boleta.idlib = mae_libros.id) ON mae_documento.id = pla_boleta.iddoc) ON pla_empleados.id = pla_boleta.idemp) ON mae_moneda.id = pla_boleta.idmon) LEFT JOIN con_tc ON pla_boleta.fchdoc = con_tc.fecha " _
                + vbCr + " ) LEFT JOIN ( " & nSQLDiario & "" _
                + vbCr + " ) AS xpag ON pla_boleta.id = xpag.iddoc " _
                + vbCr + " WHERE (((pla_boleta.iddoc)<>7)) " & nSQLWhere _
                + vbCr + " ORDER BY IIf(pla_boleta.numreg Is Null Or pla_boleta.numreg='',mae_libros.codsun,Left(pla_boleta.numreg,2) & mae_libros.codsun & Right(pla_boleta.numreg,4)), pla_boleta.fchdoc;"
            
        Case 37 '--Letras
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "let_letra")
            nSQLWhere = Replace(nSQLWhere, "con_diario.ridper", "let_letra.idclipro")
            nSQLWhere = Replace(nSQLWhere, "fchdoc", "fchemi")

            nSQL = "SELECT let_letradet.corr AS iddoc, let_letra.tipdoc, mae_cliente.numruc, mae_cliente.nombre AS nombre, IIf(let_letra.numreg Is Null Or let_letra.numreg='',mae_libros.codsun,Left(let_letra.numreg,2) & mae_libros.codsun & Right(let_letra.numreg,4)) AS registro, mae_documento.abrev, " _
                + vbCr + " let_letra.ano & ' ' & let_letradet.numdoc & ' ' & let_letradet.numser AS numdoc2, let_letra.fchemi AS fchdoc, mae_moneda.simbolo, " _
                + vbCr + " IIf(let_letra.tc Is Null Or let_letra.tc=0,con_tc.impven,let_letra.tc) AS tipcam, let_letra.idmon, let_letradet.implet AS imptotal, let_letradet.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(let_letra.idmon=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(let_letra.idmon=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, " _
                + vbCr + " let_letra.glosa ,let_letra.numerodocref as docref, " _
                + vbCr + " xtotsol, xtotdol, IIf(let_letra.idmon=1,imptotal-xtotsol,imptotal-xtotdol) AS saldo,xregistro, xnombre,xabrev,xnumdoc,xfchcancel,xglosa ,xcanreg , xsimbolo " _
                + vbCr + " FROM ( (((mae_moneda RIGHT JOIN (mae_libros RIGHT JOIN (mae_documento RIGHT JOIN let_letra ON mae_documento.id = let_letra.tipdoc) ON mae_libros.id = let_letra.idlib) ON mae_moneda.id = let_letra.idmon) INNER JOIN let_letradet ON let_letra.id = let_letradet.idlet) LEFT JOIN mae_cliente ON let_letra.idclipro = mae_cliente.id) LEFT JOIN con_tc ON let_letra.fchemi = con_tc.fecha " _
                + vbCr + " ) LEFT JOIN ( " & nSQLDiario & "" _
                + vbCr + " ) AS xpag ON let_letradet.corr = xpag.iddoc " _
                + vbCr + " WHERE (((IIf(let_letra.numreg='000001',let_letradet.imptotori,let_letradet.implet))<>0)) " & nSQLWhere _
                + vbCr + " ORDER BY IIf(let_letra.numreg Is Null Or let_letra.numreg='',mae_libros.codsun,Left(let_letra.numreg,2) & mae_libros.codsun & Right(let_letra.numreg,4)), let_letra.fchemi "
        
        Case 40 '--Honorarios
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "com_honorarios")
            nSQLWhere = Replace(nSQLWhere, "con_diario.ridper", "com_honorarios.idpro")
            
            nSQL = "SELECT  com_honorarios.id as iddoc,com_honorarios.tipdoc, mae_prov.numruc, mae_prov.nombre AS nombre, IIf(com_honorarios.numreg Is Null Or com_honorarios.numreg='',mae_libros.codsun,Left(com_honorarios.numreg,2) & mae_libros.codsun & Right(com_honorarios.numreg,4)) AS registro, mae_documento.abrev, " _
                + vbCr + " com_honorarios.numser+'-'+com_honorarios.numdoc AS numdoc2, com_honorarios.fchdoc, mae_moneda.simbolo, " _
                + vbCr + " IIf(com_honorarios.tc Is Null Or com_honorarios.tc=0,con_tc.impven,com_honorarios.tc) AS tipcam, com_honorarios.idmon, IIf(com_honorarios.numreg='000001',com_honorarios.imptotori,com_honorarios.imptot) AS imptotal, com_honorarios.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(com_honorarios.idmon=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(com_honorarios.idmon=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, " _
                + vbCr + " com_honorarios.glosa ,'' as docref, " _
                + vbCr + " xtotsol, xtotdol, IIf(com_honorarios.idmon=1,imptotal-xtotsol,imptotal-xtotdol) AS saldo, xregistro, xnombre,xabrev,xnumdoc,xfchcancel,xglosa ,xcanreg , xsimbolo " _
                + vbCr + " FROM ( (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN (mae_documento RIGHT JOIN (com_honorarios LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) ON mae_documento.id = com_honorarios.tipdoc) ON mae_prov.id = com_honorarios.idpro) ON mae_moneda.id = com_honorarios.idmon) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha " _
                + vbCr + " ) LEFT JOIN ( " & nSQLDiario & "" _
                + vbCr + " ) AS xpag ON com_honorarios.id = xpag.iddoc " _
                + vbCr + " WHERE (((com_honorarios.tipdoc)<>7) AND ((IIf(com_honorarios.numreg='000001',com_honorarios.imptotori,com_honorarios.imptot))<>0)) " & nSQLWhere _
                + vbCr + " ORDER BY IIf(com_honorarios.numreg Is Null Or com_honorarios.numreg='',mae_libros.codsun,Left(com_honorarios.numreg,2) & mae_libros.codsun & Right(com_honorarios.numreg,4)), com_honorarios.fchdoc"
    
    Case 41 '--Lgd
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "vta_gastodebito")
            nSQLWhere = Replace(nSQLWhere, "con_diario.ridper", "vta_gastodebito.idcli")
            
            nSQL = "SELECT  vta_gastodebito.id as iddoc,vta_gastodebito.tipdoc, mae_cliente.numruc, mae_cliente.nombre AS nombre, IIf(vta_gastodebito.numreg Is Null Or vta_gastodebito.numreg='',mae_libros.codsun,Left(vta_gastodebito.numreg,2) & mae_libros.codsun & Right(vta_gastodebito.numreg,4)) AS registro, mae_documento.abrev, " _
                + vbCr + " vta_gastodebito.numser+'-'+vta_gastodebito.numdoc AS numdoc2, vta_gastodebito.fchemi as fchdoc, mae_moneda.simbolo, " _
                + vbCr + " IIf(vta_gastodebito.tc Is Null Or vta_gastodebito.tc=0,con_tc.impven,vta_gastodebito.tc) AS tipcam, vta_gastodebito.idmon, vta_gastodebito.imptot AS imptotal, vta_gastodebito.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(vta_gastodebito.idmon=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(vta_gastodebito.idmon=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, " _
                + vbCr + " vta_gastodebito.glosa ,vta_gastodebito.numerodocref as docref,  " _
                + vbCr + " xtotsol, xtotdol, IIf(vta_gastodebito.idmon=1,imptotal-xtotsol,imptotal-xtotdol) AS saldo, xregistro, xnombre,xabrev,xnumdoc,xfchcancel,xglosa ,xcanreg , xsimbolo " _
                + vbCr + " FROM ( (mae_moneda RIGHT JOIN (mae_cliente RIGHT JOIN (mae_documento RIGHT JOIN (vta_gastodebito LEFT JOIN mae_libros ON vta_gastodebito.idlib = mae_libros.id) ON mae_documento.id = vta_gastodebito.tipdoc) ON mae_cliente.id = vta_gastodebito.idcli) ON mae_moneda.id = vta_gastodebito.idmon) LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha " _
                + vbCr + " ) LEFT JOIN ( " & nSQLDiario & "" _
                + vbCr + " ) AS xpag ON vta_gastodebito.id = xpag.iddoc " _
                + vbCr + " WHERE (((vta_gastodebito.tipdoc)<>7) AND ((IIf(vta_gastodebito.numreg='000001',vta_gastodebito.imptot,vta_gastodebito.imptot))<>0)) " & nSQLWhere _
                + vbCr + " ORDER BY IIf(vta_gastodebito.numreg Is Null Or vta_gastodebito.numreg='',mae_libros.codsun,Left(vta_gastodebito.numreg,2) & mae_libros.codsun & Right(vta_gastodebito.numreg,4)), vta_gastodebito.fchemi "
    
    
    Case 42 '--Planilla Letras
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "let_planilla")
            nSQLWhere = Replace(nSQLWhere, "con_diario.ridper", "mae_bancos.id")
            nSQLWhere = Replace(nSQLWhere, "fchdoc", "fchemi")
    
            nSQL = "SELECT let_planilla.id as iddoc, let_planilla.tipdoc, mae_bancos.numruc, mae_bancos.descripcion AS nombre, IIf(let_planilla.numreg Is Null Or let_planilla.numreg='',mae_libros.codsun,Left(let_planilla.numreg,2) & mae_libros.codsun & Right(let_planilla.numreg,4)) AS registro, mae_documento.abrev, " _
                + vbCr + " IIf(let_planilla.numser is null,'',let_planilla.numser & '-') & let_planilla.numdoc AS numdoc2, let_planilla.fchemi AS fchdoc, mae_moneda.simbolo, " _
                + vbCr + " IIf(let_planilla.tc Is Null Or let_planilla.tc=0,con_tc.impven,let_planilla.tc) AS tipcam, let_planilla.idmon, IIf(let_planilla.numreg='000001',let_planilla.imptot,let_planilla.imptot) AS imptotal, let_planilla.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(let_planilla.idmon=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(let_planilla.idmon=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, " _
                + vbCr + " let_planilla.glosa ,'' as docref, " _
                + vbCr + " xtotsol, xtotdol, IIf([let_planilla].[idmon]=1,[imptotal]-xtotsol,[imptotal]-xtotdol) AS saldo, xregistro, xnombre,xabrev,xnumdoc,xfchcancel,xglosa ,xcanreg , xsimbolo " _
                + vbCr + " FROM ( (((mae_bancos RIGHT JOIN (mae_banconumcta RIGHT JOIN (mae_documento RIGHT JOIN let_planilla ON mae_documento.id = let_planilla.tipdoc) ON mae_banconumcta.id = let_planilla.idbcocta) ON mae_bancos.id = mae_banconumcta.idban) LEFT JOIN mae_libros ON let_planilla.idlib = mae_libros.id) LEFT JOIN con_tc ON let_planilla.fchemi = con_tc.fecha) LEFT JOIN mae_moneda ON let_planilla.idmon = mae_moneda.id " _
                + vbCr + " ) LEFT JOIN ( " & nSQLDiario & "" _
                + vbCr + " ) AS xpag ON let_planilla.id = xpag.iddoc " _
                + vbCr + " WHERE (((IIf(let_planilla.numreg='000001',let_planilla.imptot,let_planilla.imptot))<>0)) " & nSQLWhere _
                + vbCr + " ORDER BY IIf(let_planilla.numreg Is Null Or let_planilla.numreg='',mae_libros.codsun,Left(let_planilla.numreg,2) & mae_libros.codsun & Right(let_planilla.numreg,4)), let_planilla.fchemi "
    
    End Select

    Dim xFila As Long
    
    RST_Busq xRstLineal, nSQL, xCon
    
    xFila = Fg1.FixedRows
    If xRstLineal.State = 0 Then GoTo Salir
    If xRstLineal.RecordCount = 0 Then GoTo Salir
    
    '-----------------
    fraBarra.Visible = True
    fraBarra.Left = 2798
    fraBarra.Top = 2925
    
    ProgressBar1.Max = 1
    ProgressBar1.Value = 0
    
    fraBarra.Refresh
    ProgressBar1.Max = xRstLineal.RecordCount
    BAND_INTERRUMPIR = False
    '-----------------
    Fg2.Rows = Fg2.FixedRows
    Set xRstTot = Nothing
    PreparaRST
    DoEvents
    '-----------------
    Do While Not xRstLineal.EOF
        DoEvents
        
        If BAND_INTERRUMPIR = True Then GoTo Salir:
        ProgressBar1.Value = ProgressBar1.Value + 1
            
        Fg1.Rows = Fg1.Rows + 1
        
        xFila = Fg1.Rows - 1
                
        Fg1.TextMatrix(xFila, 1) = NulosC(xRstLineal("numruc"))
        Fg1.TextMatrix(xFila, 2) = NulosC(xRstLineal("nombre"))
        Fg1.TextMatrix(xFila, 3) = NulosC(xRstLineal("registro"))
        Fg1.TextMatrix(xFila, 4) = NulosC(xRstLineal("abrev"))
        Fg1.TextMatrix(xFila, 5) = NulosC(xRstLineal("numdoc2"))
        Fg1.TextMatrix(xFila, 6) = Format(NulosC(xRstLineal("fchdoc")), FORMAT_DATE)
        Fg1.TextMatrix(xFila, 7) = NulosC(xRstLineal("simbolo"))
        Fg1.TextMatrix(xFila, 8) = NulosC(xRstLineal("tipcam"))
        Fg1.TextMatrix(xFila, 9) = Format(NulosC(xRstLineal("imptotal")), FORMAT_MONTO)
        Fg1.TextMatrix(xFila, 10) = NulosC(xRstLineal("glosa"))
        
        Fg1.TextMatrix(xFila, 11) = NulosN(xRstLineal("xcanreg"))
        Fg1.TextMatrix(xFila, 12) = Format(NulosC(xRstLineal("xfchcancel")), FORMAT_DATE)
        Fg1.TextMatrix(xFila, 13) = NulosC(xRstLineal("xregistro"))
        Fg1.TextMatrix(xFila, 14) = NulosC(xRstLineal("xnombre"))
        Fg1.TextMatrix(xFila, 15) = NulosC(xRstLineal("xabrev"))
        Fg1.TextMatrix(xFila, 16) = NulosC(xRstLineal("xnumdoc"))
        Fg1.TextMatrix(xFila, 17) = NulosC(xRstLineal("xsimbolo"))
        
        If NulosN(xRstLineal("idmon")) = 1 Then
            Fg1.TextMatrix(xFila, 18) = Format(NulosN(xRstLineal("xtotsol")), FORMAT_MONTO)
        Else
            Fg1.TextMatrix(xFila, 18) = Format(NulosN(xRstLineal("xtotdol")), FORMAT_MONTO)
        End If
        Fg1.TextMatrix(xFila, 19) = NulosC(xRstLineal("xglosa"))
        
        If NulosN(xRstLineal("xcanreg")) <> 0 Then
            Fg1.TextMatrix(xFila, 20) = Format(NulosN(xRstLineal("saldo")), FORMAT_MONTO)
        Else
            Fg1.TextMatrix(xFila, 20) = Format(NulosC(xRstLineal("imptotal")), FORMAT_MONTO)
        End If
        xImpSaldo = NulosN(Fg1.TextMatrix(xFila, 20))
        
        Fg1.TextMatrix(xFila, 21) = NulosC(xRstLineal("iddoc"))
        Fg1.TextMatrix(xFila, 22) = NulosC(xRstLineal("idmon"))
        Fg1.TextMatrix(xFila, 23) = NulosC(xRstLineal("docref"))
        
        If xImpSaldo <> NulosN(xRstLineal("impsal")) Then
        
            '--Actualizar saldos a documento
            Select Case xIdLibro
                Case 1   '--Compras
                    xCon.Execute "Update com_compras set com_compras.impsal=" & xImpSaldo & " where com_compras.id = " & NulosC(xRstLineal("iddoc"))
                Case 4   '--Percepcion
                    xCon.Execute "Update con_percepcion set con_percepcion.impsal=" & xImpSaldo & " where con_percepcion.id = " & NulosC(xRstLineal("iddoc"))
                Case 40  '--Honorarios
                    xCon.Execute "Update com_honorarios set com_honorarios.impsal=" & xImpSaldo & " where com_honorarios.id = " & NulosC(xRstLineal("iddoc"))
                Case 2   '--Ventas
                    xCon.Execute "Update vta_ventas set vta_ventas.impsal=" & xImpSaldo & " where vta_ventas.id = " & NulosC(xRstLineal("iddoc"))
                Case 9  '--Boleta Pago
                    xCon.Execute "Update pla_boleta set pla_boleta.impsal=" & xImpSaldo & " where pla_boleta.id = " & NulosC(xRstLineal("iddoc"))
                Case 37  '--Letras
                    xCon.Execute "Update let_letradet set let_letradet.impsal=" & xImpSaldo & " where let_letradet.corr = " & NulosC(xRstLineal("iddoc"))
                Case 41  '--Lgd, Lgc
                    xCon.Execute "Update vta_gastodebito set vta_gastodebito.impsal=" & xImpSaldo & " where vta_gastodebito.id = " & NulosC(xRstLineal("iddoc"))
                Case 42  '--Planilla letras
                    xCon.Execute "Update let_planilla set let_planilla.impsal=" & xImpSaldo & " where let_planilla.id = " & NulosC(xRstLineal("iddoc"))
                Case 999 '--Reembolsables
                    xCon.Execute "Update com_reembolsables set com_reembolsables.impsal=" & xImpSaldo & " where com_reembolsables.id = " & NulosC(xRstLineal("iddoc"))
            End Select
        
        End If
        
        '--pintar las celdas
        If NulosN(xRstLineal("xcanreg")) = 0 Then
            '--Pendientes de agregar operaciones
            
        ElseIf NulosN(Format(xRstLineal("saldo"), FORMAT_MONTO)) = 0 Then
            '--Documentos Cancelados
            GRID_COLOR_FONDO Fg1, xFila, 1, xFila, Fg1.Cols - 1, &HA4FFA4
            
        ElseIf NulosN(Format(xRstLineal("saldo"), FORMAT_MONTO)) < 0 Then
            '--Documentos Observados
            GRID_COLOR_FONDO Fg1, xFila, 1, xFila, Fg1.Cols - 1, &H8C8CFF
            
        Else
            '--Documentos Pendientes
            GRID_COLOR_FONDO Fg1, xFila, 1, xFila, Fg1.Cols - 1, vbYellow '&H9BFFFF
        End If
        
        xEliminaReg = False
        
        If OptPen.Value = True Then
            If NulosN(xRstLineal("xcanreg")) <> 0 And NulosN(Format(xRstLineal("saldo"), FORMAT_MONTO)) = 0 Then
                Fg1.Rows = Fg1.Rows - 1
                xEliminaReg = True
            End If
        ElseIf OptCan.Value = True Then
            If NulosN(xRstLineal("xcanreg")) = 0 Or NulosN(Format(xRstLineal("saldo"), FORMAT_MONTO)) <> 0 Then
                Fg1.Rows = Fg1.Rows - 1
                xEliminaReg = True
            End If
        Else
            
        End If
        If xEliminaReg = False Then
            Acumular NulosC(xRstLineal("numruc")), NulosC(xRstLineal("nombre")), _
                     IIf(NulosN(xRstLineal("idmon")) = 1, NulosN(xRstLineal("imptotal")), 0), _
                     IIf(NulosN(xRstLineal("idmon")) = 2, NulosN(xRstLineal("imptotal")), 0), _
                     IIf(NulosN(xRstLineal("idmon")) = 1, NulosN(xRstLineal("xtotsol")), 0), _
                     IIf(NulosN(xRstLineal("idmon")) = 2, NulosN(xRstLineal("xtotdol")), 0)
        End If
        
        xRstLineal.MoveNext
    Loop
    
    '----------------
    '-- cargar datos del resumen
    xRstTot.Filter = ""
    If xRstTot.RecordCount <> 0 Then xRstTot.MoveFirst
    Do While Not xRstTot.EOF
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(xRstTot("ruc"))
        Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosC(xRstTot("nombre"))
        Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(NulosN(xRstTot("impmn")), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(NulosN(xRstTot("impme")), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(NulosN(xRstTot("pimpmn")), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 6) = Format(NulosN(xRstTot("pimpme")), FORMAT_MONTO)
        
        Fg2.TextMatrix(Fg2.Rows - 1, 7) = Format(NulosN(xRstTot("impmn")) - NulosN(xRstTot("pimpmn")), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 8) = Format(NulosN(xRstTot("impme")) - NulosN(xRstTot("pimpme")), FORMAT_MONTO)
        
        xRstTot.MoveNext
    Loop
    
    If Fg2.Rows > Fg2.FixedRows Then
        '--Ordenar
        GRID_ORDENAR Fg2, Fg2.FixedRows, 2
        
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(Fg2.Rows - 1, 2) = "TOTALES "
        Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(NulosN(GRID_SUMAR_COL(Fg2, 3)), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(NulosN(GRID_SUMAR_COL(Fg2, 4)), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(NulosN(GRID_SUMAR_COL(Fg2, 5)), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 6) = Format(NulosN(GRID_SUMAR_COL(Fg2, 6)), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 7) = Format(NulosN(GRID_SUMAR_COL(Fg2, 7)), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 8) = Format(NulosN(GRID_SUMAR_COL(Fg2, 8)), FORMAT_MONTO)

        With Fg2
            .Cell(flexcpForeColor, Fg2.Rows - 1, 1, Fg2.Rows - 1, Fg2.Cols - 1) = &H800000
            .Select Fg2.Rows - 1, 1, Fg2.Rows - 1, Fg2.Cols - 1
            .FillStyle = flexFillRepeat
            .CellFontBold = True
        End With
        
        '--Ajustar las columnas
        Fg2.AutoSizeMode = flexAutoSizeColWidth
        Fg2.AutoSize 3
        Fg2.AutoSize 4
        Fg2.AutoSize 5
        Fg2.AutoSize 6
        Fg2.AutoSize 7
        Fg2.AutoSize 8
        
    End If
    
Salir:

    Set xRstLineal = Nothing
    fraBarra.Visible = False
    BAND_INTERRUMPIR = False
End Sub


Private Sub VerDatosDetalle()
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim xIdDoc As Long
    Dim xBase As String
    Dim nSQLDocNCBancos As String
    
    If BAND_INTERRUMPIR = True Then Exit Sub
    
    FraDetalle.Visible = True
    FraDetalle.Left = 1650
    FraDetalle.Top = 2070
    '--ocultar columna correlativo
    Fg3.ColWidth(1) = 0
    
    '--Obtener codigo del documento
    xIdDoc = NulosN(Fg1.TextMatrix(Fg1.Row, 21))
    
    If xIdLibro = 1 Or xIdLibro = 4 Or xIdLibro = 9 Or xIdLibro = 40 Or xIdLibro = 999 Then
        '--compras, honorarios, reembolsables, boleta pago
        xBase = vbCr & " IIf(con_diario.tipmov =1, IIf(con_diario.tipo =1,IIf(con_diario.rtipdoc=7,-1,1),IIf(con_diario.rtipdoc=7,1,-1)) , IIf(con_diario.tipo in (1),IIf(con_diario.rtipdoc=7,1,-1),1) ) as xbase, "
    Else
        '--ventas, lgd, letras, planilla letras
        xBase = " IIf(con_diario.tipmov =1, IIf(con_diario.tipo =1,IIf(con_diario.rtipdoc=7,1,-1),1), IIf(con_diario.tipo in (0,1),1,IIf(con_diario.rtipdoc=7,1,-1)) ) as xbase, "
    End If
        
    nSQL = " SELECT con_diario.rregistro, Format([con_diario].[idmes],'00') & [mae_libros].[codsun] & Format([con_diario].[numasi],'0000') AS registro, mae_libros.descripcion AS libro, IIf([con_diario].[ridtipper2]=5,[mae_bancos].[abrev],IIf([con_diario].[ridtipper2]=2,[mae_cliente].[nombre],IIf([con_diario].[ridtipper2]=1,[mae_prov].[nombre],''))) AS razonsocial, tes_documentos.abrev, con_diario.rnumerodoc2 AS numdoc, con_diario.rfchope2 AS fchemi, mae_moneda.simbolo, IIf([con_diario].[aplicatc]=0,[con_tc].[impven],[con_diario].[tc]) AS tipcam, " _
        + vbCr + " con_diario.tipmov, con_diario.tipo, " & xBase _
        + vbCr + " IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS imptotal, " _
        + vbCr + " IIf(con_diario.idmon=1,imptotal,imptotal*tipcam) * xbase  AS imptotsol, " _
        + vbCr + " IIf(con_diario.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) * xbase AS imptotdol, " _
        + vbCr + " con_diario.ridper, con_diario.rnumerodoc AS numdoc2, con_diario.rglosaope, con_diario.iddoc, con_diario.idmon " _
        + vbCr + " FROM ((((((con_diario LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN tes_documentos ON con_diario.rtipdoc2 = tes_documentos.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN mae_bancos ON con_diario.ridper2 = mae_bancos.id) LEFT JOIN mae_cliente ON con_diario.ridper2 = mae_cliente.id) LEFT JOIN mae_prov ON con_diario.ridper2 = mae_prov.id " _
        + vbCr + " WHERE (((con_diario.idlib) In (" & xIdLibroRef & ")) AND ((con_diario.ridlib)=" & xIdLibro & ") AND ((con_diario.iddoc)=" & xIdDoc & ")) "
    
    If xIdLibro = 1 Then
    
        '--Verificar si hay documentos de NC que fueron registrados en Tesoreria Ingresos - Egresos
        nSQLDocNCBancos = BuscarNCBancos()
        If nSQLDocNCBancos <> "" Then
            nSQLDocNCBancos = " and com_compras.id not in (" & nSQLDocNCBancos & ")"
        End If
        
        If Fg1.TextMatrix(Fg1.Row, 4) <> "NC" Then
            '--Cancelacion de compras con nota de credito excepto nc que se registran en tesoreria
            nSQL = nSQL _
                + vbCr + " UNION " _
                + vbCr + " SELECT Left(com_compras_1.numreg,2) & mae_libros_1.codsun & Right(com_compras_1.numreg,4) AS rregistro, Mid(com_compras!numreg,1,2) & mae_libros!codsun & Mid(com_compras.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_prov.nombre AS razonsocial, mae_documento.abrev, " _
                + vbCr + " com_compras.numser & '-' & com_compras.numdoc AS numdoc, com_compras.fchdoc AS fchemi, mae_moneda.simbolo, " _
                + vbCr + " IIf(com_compras_1.tc=0,con_tc.impven,com_compras_1.tc) AS tipcam, 2 AS tipmov, 1 AS tipo, 1 AS xbase, com_compras.imptot AS imptotal, " _
                + vbCr + " IIf(com_compras.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, " _
                + vbCr + " IIf(com_compras.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) AS imptotdol, " _
                + vbCr + " com_compras.idpro AS ridper, com_compras_1.numser & '-' & com_compras_1.numdoc AS numdoc2, com_compras.glosa AS glosaope, com_compras_1.id AS iddoc, com_compras.idmon " _
                + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((((com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) INNER JOIN com_compras AS com_compras_1 ON com_compras.iddocref = com_compras_1.id)  " _
                + vbCr + " LEFT JOIN mae_libros AS mae_libros_1 ON com_compras_1.idlib = mae_libros_1.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) LEFT JOIN mae_prov ON com_compras.idpro = mae_prov.id " _
                + vbCr + " WHERE (com_compras.iddocref<>0 ) and com_compras_1.id = " & xIdDoc & nSQLDocNCBancos
    
        Else
            '--Cancelacion de notas de credito con compras excepto nc que se registran en tesoreria
            nSQLDocNCBancos = Replace(nSQLDocNCBancos, "com_compras", "com_compras_1")
    
            nSQL = nSQL _
                + vbCr + " UNION " _
                + vbCr + " SELECT Left(com_compras_1.numreg,2) & mae_libros_1.codsun & Right(com_compras_1.numreg,4) AS rregistro, Mid(com_compras!numreg,1,2) & mae_libros!codsun & Mid(com_compras.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_prov.nombre AS razonsocial, mae_documento.abrev, " _
                + vbCr + " com_compras.numser & '-' & com_compras.numdoc AS numdoc, com_compras.fchdoc AS fchemi, mae_moneda.simbolo, " _
                + vbCr + " IIf(com_compras_1.tc=0,con_tc.impven,com_compras_1.tc) AS tipcam, 2 AS tipmov, 1 AS tipo, 1 AS xbase, com_compras_1.imptot AS imptotal, " _
                + vbCr + " IIf(com_compras.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, " _
                + vbCr + " IIf(com_compras.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) AS imptotdol, " _
                + vbCr + " com_compras.idpro AS ridper, com_compras_1.numser & '-' & com_compras_1.numdoc AS numdoc2, com_compras.glosa AS glosaope, com_compras_1.id AS iddoc, com_compras.idmon " _
                + vbCr + " FROM (com_compras AS com_compras_1 LEFT JOIN mae_libros AS mae_libros_1 ON com_compras_1.idlib = mae_libros_1.id) INNER JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (((com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) " _
                + vbCr + " LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) LEFT JOIN mae_prov ON com_compras.idpro = mae_prov.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON com_compras_1.iddocref = com_compras.id " _
                + vbCr + " WHERE (com_compras_1.iddocref<>0 ) and com_compras_1.id = " & xIdDoc & nSQLDocNCBancos
        End If
        '--------------------------------------------------
    
'        nSQL = nSQL _
            + vbCr + " UNION " _
            + vbCr + " SELECT Left(com_compras_1.numreg,2) & mae_libros_1.codsun & Right(com_compras_1.numreg,4) AS rregistro, Mid(com_compras.numreg,1,2) & mae_libros.codsun & Mid(com_compras.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_prov.nombre AS razonsocial, mae_documento.abrev, " _
            + vbCr + " com_compras.numser & '-' & com_compras.numdoc AS numdoc, com_compras.fchdoc AS fchemi, mae_moneda.simbolo, " _
            + vbCr + " IIf(com_compras_1.tc=0,con_tc.impven,com_compras_1.tc) AS tipcam, 2 AS tipmov, 1 AS tipo, 1 AS xbase, com_compras.imptot AS imptotal, " _
            + vbCr + " IIf(com_compras.idmon=1,com_compras.imptot,com_compras.imptot*tipcam) AS imptotsol, IIf(com_compras.idmon=2,com_compras.imptot,IIf(tipcam=0,0,com_compras.imptot/tipcam)) AS imptotdol, " _
            + vbCr + " com_compras.idpro AS ridper, com_compras_1.numser & '-' & com_compras_1.numdoc AS numdoc2, com_compras.glosa AS glosaope, com_compras_1.id AS iddoc, com_compras.idmon,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc " _
            + vbCr + " FROM (((mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((((com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) INNER JOIN com_compras AS com_compras_1 ON com_compras.iddocref = com_compras_1.id) LEFT JOIN mae_libros AS mae_libros_1 ON com_compras_1.idlib = mae_libros_1.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) LEFT JOIN (con_planctas RIGHT JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (com_compras.tipdoc = mae_documentocta.iddoc) AND (com_compras.idmon = mae_documentocta.idmon)) " _
            + vbCr + " LEFT JOIN " _
            + vbCr + " (SELECT tes_cajaorigendet.iddoc FROM tes_cajaorigendet INNER JOIN com_compras ON tes_cajaorigendet.iddoc = com_compras.id WHERE (((tes_cajaorigendet.idmod)=1) AND ((com_compras.tipdoc)=7)) " _
            + vbCr + "  ) AS tes ON com_compras.id = tes.iddoc) LEFT JOIN mae_prov ON com_compras.idpro = mae_prov.id " _
            + vbCr + " WHERE (((com_compras.iddocref) Is Not Null And (com_compras.iddocref)<>0) AND ((tes.iddoc) Is Null)) and mae_documentocta.tipope =0 and com_compras_1.id = " & xIdDoc
       
    
    ElseIf xIdLibro = 2 Then
        '--Verificar si hay documentos de NC que fueron registrados en Tesoreria Ingresos - Egresos
        nSQLDocNCBancos = BuscarNCBancos()
        If nSQLDocNCBancos <> "" Then
            nSQLDocNCBancos = " and vta_ventas.id not in (" & nSQLDocNCBancos & ")"
        End If
        
        If Fg1.TextMatrix(Fg1.Row, 4) <> "NC" Then
            '--Cancelacion de compras con nota de credito excepto nc que se registran en tesoreria
            nSQL = nSQL _
                + vbCr + " UNION " _
                + vbCr + " SELECT Left(vta_ventas_1.numreg,2) & mae_libros_1.codsun & Right(vta_ventas_1.numreg,4) AS rregistro, Mid(vta_ventas.numreg,1,2) & mae_libros.codsun & Mid(vta_ventas.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_cliente.nombre AS razonsocial, mae_documento.abrev, " _
                + vbCr + " vta_ventas.numser & '-' & vta_ventas.numdoc AS numdoc, vta_ventas.fchdoc AS fchemi, mae_moneda.simbolo, " _
                + vbCr + " IIf(vta_ventas_1.tc=0,con_tc.impven,vta_ventas_1.tc) AS tipcam, 2 AS tipmov, 1 AS tipo, 1 AS xbase, vta_ventas.imptotdoc AS imptotal, " _
                + vbCr + " IIf(vta_ventas.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, " _
                + vbCr + " IIf(vta_ventas.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) AS imptotdol, " _
                + vbCr + " vta_ventas.idcli AS ridper, vta_ventas_1.numser & '-' & vta_ventas_1.numdoc AS numdoc2, vta_ventas.glosa AS glosaope, vta_ventas_1.id AS iddoc, vta_ventas.idmon " _
                + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((((vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) INNER JOIN vta_ventas AS vta_ventas_1 ON vta_ventas.iddocref = vta_ventas_1.id)  " _
                + vbCr + " LEFT JOIN mae_libros AS mae_libros_1 ON vta_ventas_1.idlib = mae_libros_1.id) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id " _
                + vbCr + " WHERE (vta_ventas.iddocref<>0 ) and vta_ventas_1.id = " & xIdDoc & nSQLDocNCBancos
    
        Else
            '--Cancelacion de notas de credito con compras excepto nc que se registran en tesoreria
            nSQLDocNCBancos = Replace(nSQLDocNCBancos, "vta_ventas", "vta_ventas_1")
    
            nSQL = nSQL _
                + vbCr + " UNION " _
                + vbCr + " SELECT Left(vta_ventas_1.numreg,2) & mae_libros_1.codsun & Right(vta_ventas_1.numreg,4) AS rregistro, Mid(vta_ventas.numreg,1,2) & mae_libros.codsun & Mid(vta_ventas.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_cliente.nombre AS razonsocial, mae_documento.abrev, " _
                + vbCr + " vta_ventas.numser & '-' & vta_ventas.numdoc AS numdoc, vta_ventas.fchdoc AS fchemi, mae_moneda.simbolo, " _
                + vbCr + " IIf(vta_ventas_1.tc=0,con_tc.impven,vta_ventas_1.tc) AS tipcam, 2 AS tipmov, 1 AS tipo, 1 AS xbase, vta_ventas_1.imptotdoc AS imptotal, " _
                + vbCr + " IIf(vta_ventas.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, " _
                + vbCr + " IIf(vta_ventas.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) AS imptotdol, " _
                + vbCr + " vta_ventas.idcli AS ridper, vta_ventas_1.numser & '-' & vta_ventas_1.numdoc AS numdoc2, vta_ventas.glosa AS glosaope, vta_ventas_1.id AS iddoc, vta_ventas.idmon " _
                + vbCr + " FROM (vta_ventas AS vta_ventas_1 LEFT JOIN mae_libros AS mae_libros_1 ON vta_ventas_1.idlib = mae_libros_1.id) INNER JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (((vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) " _
                + vbCr + " LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON vta_ventas_1.iddocref = vta_ventas.id " _
                + vbCr + " WHERE (vta_ventas_1.iddocref<>0 ) and vta_ventas_1.id = " & xIdDoc & nSQLDocNCBancos
        End If
    
'        nSQL = nSQL _
            + vbCr + " UNION " _
            + vbCr + " SELECT Left(vta_ventas_1.numreg,2) & mae_libros_1.codsun & Right(vta_ventas_1.numreg,4) AS rregistro, Mid(vta_ventas.numreg,1,2) & mae_libros.codsun & Mid(vta_ventas.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_prov.nombre AS razonsocial, mae_documento.abrev, " _
            + vbCr + " vta_ventas.numser & '-' & vta_ventas.numdoc AS numdoc, vta_ventas.fchdoc AS fchemi, mae_moneda.simbolo, " _
            + vbCr + " IIf(vta_ventas_1.tc=0,con_tc.impven,vta_ventas_1.tc) AS tipcam, 2 AS tipmov, 1 AS tipo, 1 AS xbase, vta_ventas.imptotdoc AS imptotal, " _
            + vbCr + " IIf(vta_ventas.idmon=1,vta_ventas.imptotdoc,vta_ventas.imptotdoc*tipcam) AS imptotsol, IIf(vta_ventas.idmon=2,vta_ventas.imptotdoc,IIf(tipcam=0,0,vta_ventas.imptotdoc/tipcam)) AS imptotdol, " _
            + vbCr + " vta_ventas.idcli AS ridper, vta_ventas_1.numser & '-' & vta_ventas_1.numdoc AS numdoc2, vta_ventas.glosa AS glosaope, vta_ventas_1.id AS iddoc, vta_ventas.idmon,con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc " _
            + vbCr + " FROM (((mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((((vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) INNER JOIN vta_ventas AS vta_ventas_1 ON vta_ventas.iddocref = vta_ventas_1.id) LEFT JOIN mae_libros AS mae_libros_1 ON vta_ventas_1.idlib = mae_libros_1.id) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) LEFT JOIN (con_planctas RIGHT JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (vta_ventas.tipdoc = mae_documentocta.iddoc) AND (vta_ventas.idmon = mae_documentocta.idmon)) " _
            + vbCr + " LEFT JOIN " _
            + vbCr + " (SELECT tes_cajaorigendet.iddoc FROM tes_cajaorigendet INNER JOIN vta_ventas ON tes_cajaorigendet.iddoc = vta_ventas.id WHERE (((tes_cajaorigendet.idmod)=2) AND ((vta_ventas.tipdoc)=7)) " _
            + vbCr + "  ) AS tes ON vta_ventas.id = tes.iddoc) LEFT JOIN mae_prov ON vta_ventas.idcli = mae_prov.id " _
            + vbCr + " WHERE vta_ventas.anulado=0 and (((vta_ventas.iddocref) Is Not Null And (vta_ventas.iddocref)<>0) AND ((tes.iddoc) Is Null)) and mae_documentocta.tipope =-1 and vta_ventas_1.id = " & xIdDoc
   
   
    End If
    
    RST_Busq xRs, nSQL, xCon
    If xRs.State = 0 Then
        Set xRs = Nothing
        FraDetalle.Visible = False
        Exit Sub
    End If
    
    '--------------------------------------------------------
    Fg3.Rows = 1
    DoEvents
    '--------------------------------------------------------
    '
    '--Mostrando datos del documento
    lblDetNombre1.Caption = Fg1.TextMatrix(Fg1.Row, 2)
    lblDetNumDoc1.Caption = Fg1.TextMatrix(Fg1.Row, 4) & " / " & Fg1.TextMatrix(Fg1.Row, 5)
    lblDetFchEmi1.Caption = Fg1.TextMatrix(Fg1.Row, 6)
    lblDetImp1.Caption = Fg1.TextMatrix(Fg1.Row, 7) & " " & Fg1.TextMatrix(Fg1.Row, 9)
    lblDetTC1.Caption = Fg1.TextMatrix(Fg1.Row, 8)
    
    '--Cargando detalle documentos
    If xRs.RecordCount <> 0 Then
        xRs.Sort = "fchemi asc, registro asc"
        Do While Not xRs.EOF
            Fg3.Rows = Fg3.Rows + 1
            Fg3.TextMatrix(Fg3.Rows - 1, 1) = Fg3.Rows - 1
            Fg3.TextMatrix(Fg3.Rows - 1, 2) = NulosC(xRs("registro"))
            Fg3.TextMatrix(Fg3.Rows - 1, 3) = NulosC(xRs("libro"))
            Fg3.TextMatrix(Fg3.Rows - 1, 4) = NulosC(xRs("razonsocial"))
            Fg3.TextMatrix(Fg3.Rows - 1, 5) = NulosC(xRs("abrev"))
            Fg3.TextMatrix(Fg3.Rows - 1, 6) = NulosC(xRs("numdoc"))
            Fg3.TextMatrix(Fg3.Rows - 1, 7) = Format(NulosC(xRs("fchemi")), FORMAT_DATE)
            Fg3.TextMatrix(Fg3.Rows - 1, 8) = NulosC(xRs("simbolo"))
            
            Fg3.TextMatrix(Fg3.Rows - 1, 9) = Format(NulosN(xRs("tipcam")), "0.000")
            Fg3.TextMatrix(Fg3.Rows - 1, 10) = Format(NulosN(xRs("imptotsol")), FORMAT_MONTO)
            Fg3.TextMatrix(Fg3.Rows - 1, 11) = Format(NulosN(xRs("imptotdol")), FORMAT_MONTO)
            Fg3.TextMatrix(Fg3.Rows - 1, 12) = NulosC(xRs("rglosaope"))
                        
            xRs.MoveNext
        Loop

    End If
    
    Set xRs = Nothing

    '--verificar si hay registros en detalle
    If Fg3.Rows = Fg3.FixedRows Then
        MsgBox "No hay datos en el detalle", vbInformation, xTitulo
        FraDetalle.Visible = False
        Exit Sub
    End If
    
    '--totalizando
    Fg3.Rows = Fg3.Rows + 1
    Fg3.TextMatrix(Fg3.Rows - 1, 7) = "Totales: "
    Fg3.TextMatrix(Fg3.Rows - 1, 10) = Format(GRID_SUMAR_COL(Fg3, 10), FORMAT_MONTO)
    Fg3.TextMatrix(Fg3.Rows - 1, 11) = Format(GRID_SUMAR_COL(Fg3, 11), FORMAT_MONTO)
    
    GRID_COLOR_FONDO Fg3, Fg3.Rows - 1, 1, Fg3.Rows - 1, Fg3.Cols - 1
    
    '--Mostrando detalle
    
    Fg3.Rows = Fg3.Rows + 1
    Fg3.TextMatrix(Fg3.Rows - 1, 7) = "Saldo: "
    
    If Fg1.TextMatrix(Fg1.Row, 22) = 1 Then
        Fg3.TextMatrix(Fg3.Rows - 1, 10) = Format(NulosN(Fg1.TextMatrix(Fg1.Row, 9)) - NulosN(Fg3.TextMatrix(Fg3.Rows - 2, 10)), FORMAT_MONTO)
        Fg3.TextMatrix(Fg3.Rows - 1, 11) = " "
    Else
        Fg3.TextMatrix(Fg3.Rows - 1, 11) = Format(NulosN(Fg1.TextMatrix(Fg1.Row, 9)) - NulosN(Fg3.TextMatrix(Fg3.Rows - 2, 11)), FORMAT_MONTO)
        Fg3.TextMatrix(Fg3.Rows - 1, 10) = " "
    End If
    
    
    
    GRID_COLOR_FONDO Fg3, Fg3.Rows - 1, 1, Fg3.Rows - 1, Fg3.Cols - 1
    
    
    '--posicionar en ultima fila
    Fg3.Row = Fg3.Rows - 1
    
    CmdSalir.SetFocus
    
    DoEvents

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

Private Sub CmdSalir_Click()
    FraDetalle.Visible = False
End Sub

Private Sub pic1_Click()
    CmdSalir_Click
End Sub

Sub CargarCli4(IdCliPro)
    '===================================================================================================
    'creado: 27/12/11 Johan Castro
    'Propósito: Generar la consulta a nivel de detalle
    '
    'Entradas:  IdCliPro=Codigo del proveedor, cliente, prestador de servicio OPCIONAL=0
    '
    'Resultados: Consulta segun parametros indicados
    '
    'Modificado:
    '===================================================================================================

    
    Dim rst As New ADODB.Recordset
    Dim Rstabo As New ADODB.Recordset
    Dim A, B, xFila As Long
    Dim TotDebe, TotHaber As Double
    Dim TotGralDebe, TotGralHaber As Double
    Dim xNomPro As String '--Razon social del proveedor, cliente, prestador de servicio
    Dim Cambio As Boolean
    Dim nSQL As String
    Dim sSaldoFinal As Double '--indica el saldo final por cada documento
    Dim nSQLDocNCBancos As String '--almacenara los documentos de NC que pasan x banco
        
'    On Error GoTo error
    
    '--posicionar en vista de inicio
    TabOne2.CurrTab = 0
    
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "Seleccione una Moneda", vbExclamation, xTitulo
        TxtIdMon.SetFocus
        Exit Sub
    End If
    BAND_INTERRUMPIR = False
    pConfigurarGrilla
    
   '-----------------------------------------------
    If Fg4.Row < 1 Then Exit Sub
    xIdLibro = Fg4.TextMatrix(Fg4.Row, 2)
    Select Case xIdLibro
        Case 1, 4, 40, 9
            xIdLibroRef = "6,8,39,44"
        Case 2 '--ventas
            xIdLibroRef = "5,6,8,37,44"
        Case 37 '--Letras
            xIdLibroRef = "6,42,44"
        Case 41 '--Lgd
            xIdLibroRef = "6,41,44"
        Case 42 '--Planilla Letras
            xIdLibroRef = "6,44"
        Case 999
            If OptReem1.Value = True Then
                xIdLibroRef = "6"
            Else
                xIdLibroRef = "41"
            End If
    End Select
    '-----------------------------------------------
    
    fraBarra.Left = 2798
    fraBarra.Top = 2925
    
    TabOne1.CurrTab = 1
    
    ProgressBar1.Max = 1
    ProgressBar1.Value = 0
    fraBarra.Visible = True
    fraBarra.Refresh
    DoEvents
    
    
    Dim nSQLWhere As String '--almacenara la condicion de la consulta
    Dim nCampoMuestra As String '--indica el campo que se mostrara esta en funcion de la moneda seleccionada
    Dim nSQLAjuste  As String
    Dim nSQLApertura As String '--filtro para documentos de apertura
    Dim nSQLFecha As String '--filtro por intervalo de fechas
    Dim xOrigen As String '--Indica el origen del importe D=Debe, H=Haber
    '--para ajuste por diferencia de cambio
    nSQLAjuste = " and (con_diario.ajuste in (0, " & NulosN(TxtIdMon.Text) & ") ) "
    '------
    nSQLWhere = ""
        
    '--aplicar filtro por fecha
    If OptFch(0).Value = True Then '--x fecha de documento
        nSQLFecha = " and ( vta_ventas.fchdoc between CDate('" & TxtFchIni.Valor & "') and CDate('" & TxtFchFin.Valor & "') )"
    ElseIf OptFch(1).Value = True Then '--x fecha de registro
        nSQLFecha = " and ( vta_ventas.fchreg between CDate('" & TxtFchIni.Valor & "') and CDate('" & TxtFchFin.Valor & "') )"
    End If
    
    '--documentos de apertura
    If IdCliPro <> 0 Then
        nSQLWhere = " and vta_ventas.idcli = " & IdCliPro & " "
        If OptAperturaCon.Value = True Then nSQLApertura = " or (vta_ventas.numreg='000001' " & nSQLWhere & " ) "
    Else
        If OptAperturaCon.Value = True Then nSQLApertura = " or vta_ventas.numreg='000001' "
    End If
    If OptAperturaSin.Value = True Then nSQLApertura = " and vta_ventas.numreg<>'000001' "
    If OptAperturaSolo.Value = True Then nSQLApertura = " and vta_ventas.numreg='000001' "
    
    nSQLWhere = nSQLWhere & nSQLFecha & nSQLApertura

    Select Case xIdLibro
        Case 1 '--compras
            '--reemplazar filtro
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "com_compras")
            nSQLWhere = Replace(nSQLWhere, "idcli", "idpro")
            '------------
            nSQL = "SELECT  com_compras.tipdoc,com_compras.id,mae_prov.numruc, mae_prov!nombre AS nombre, IIf(com_compras.numreg Is Null Or com_compras.numreg='',mae_libros.codsun ,Left(com_compras.numreg,2) & mae_libros.codsun & Right(com_compras.numreg,4)) AS registro, " _
                + vbCr + " 'Compras' AS libro, mae_documento.codsun,mae_documento.abrev, iif(com_compras!numser is null or com_compras!numser ='','',com_compras!numser  +'-' ) + com_compras!numdoc AS numdoc2, com_compras.fchdoc, com_compras.fchven, mae_moneda.simbolo, " _
                + vbCr + " iif(com_compras.tc is null or com_compras.tc=0,con_tc.impven , com_compras.tc) AS tipcam,com_compras.idmon, " _
                + vbCr + " IIf(com_compras.numreg='000001',com_compras.imptotori,com_compras.imptot) AS imptotal, com_compras.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(com_compras.idmon=1, imptotal ,IIf(tipcam Is Null,0, imptotal * tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(com_compras.idmon=2, imptotal ,IIf(tipcam Is Null,0, imptotal / tipcam))) AS imptotdol, " _
                + vbCr + " com_compras.glosa as glosaope,IIf(com_compras.tipdoc<>7,'H','D') as xOrigen " _
                + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN (mae_documento RIGHT JOIN (com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_documento.id = com_compras.tipdoc) ON mae_prov.id = com_compras.idpro) ON mae_moneda.id = com_compras.idmon) " _
                + vbCr + "         LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
                + vbCr + " WHERE (IIf(com_compras.numreg='000001',com_compras.imptotori,com_compras.imptot)<>0) " & nSQLWhere & "" _
                + vbCr + " ORDER BY mae_prov!nombre, com_compras.fchdoc "
                          
'            nSQL = "SELECT  com_compras.tipdoc,com_compras.id,mae_prov.numruc, mae_prov!nombre AS nombre, IIf(com_compras.numreg Is Null Or com_compras.numreg='',mae_libros.codsun ,Left(com_compras.numreg,2) & mae_libros.codsun & Right(com_compras.numreg,4)) AS registro, " _
                + vbCr + " 'Compras' AS libro, mae_documento.codsun,mae_documento.abrev, iif(com_compras!numser is null or com_compras!numser ='','',com_compras!numser  +'-' ) + com_compras!numdoc AS numdoc2, com_compras.fchdoc, com_compras.fchven, mae_moneda.simbolo, " _
                + vbCr + " iif(com_compras.tc is null or com_compras.tc=0,con_tc.impven , com_compras.tc) AS tipcam,com_compras.idmon, " _
                + vbCr + " IIf(com_compras.numreg='000001',com_compras.imptotori,com_compras.imptot) AS imptotal, com_compras.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(com_compras.idmon=1, imptotal ,IIf(tipcam Is Null,0, imptotal * tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(com_compras.idmon=2, imptotal ,IIf(tipcam Is Null,0, imptotal / tipcam))) AS imptotdol, " _
                + vbCr + " com_compras.glosa as glosaope,IIf(com_compras.tipdoc<>7,'H','D') as xOrigen " _
                + vbCr + " FROM ( (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN (mae_documento RIGHT JOIN (com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_documento.id = com_compras.tipdoc) ON mae_prov.id = com_compras.idpro) ON mae_moneda.id = com_compras.idmon) " _
                + vbCr + "         LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
                + vbCr + "       ) LEFT JOIN " _
                + vbCr + " ( SELECT tes_cajaorigendet.iddoc FROM tes_cajaorigendet INNER JOIN com_compras ON tes_cajaorigendet.iddoc = com_compras.id WHERE (((tes_cajaorigendet.idmod)=1) AND ((com_compras.tipdoc)=7)) " _
                + vbCr + "   UNION " _
                + vbCr + "   SELECT tes_cajadestinodet.iddoc FROM tes_cajadestinodet INNER JOIN com_compras ON tes_cajadestinodet.iddoc = com_compras.id WHERE (((tes_cajadestinodet.idmod)=1) AND ((com_compras.tipdoc)=7)) " _
                + vbCr + "  ) as tes ON com_compras.id=tes.iddoc " _
                + vbCr + " WHERE ( (IIf(com_compras.numreg='000001',com_compras.imptotori,com_compras.imptot)<>0) AND " _
                + vbCr + "          (com_compras.tipdoc <>7 or com_compras.tipdoc=7 AND com_compras.iddocref=0) OR " _
                + vbCr + "          (tes.iddoc is not null) ) " & nSQLWhere & "" _
                + vbCr + " ORDER BY mae_prov!nombre, com_compras.fchdoc "
                                                    
                          
            '--percepciones ------------------------------------------------------------------------
            '--reemplazar filtro
''''            nSQLWhere = Replace(nSQLWhere, "com_compras", "con_percepcion")
''''            nSQLWhere = Replace(nSQLWhere, "idpro", "idcli")
''''            '------------
''''            nSQL = nSQL + vbCr + " Union " _
''''                + vbCr + "SELECT con_percepcion.tipdoc,con_percepcion.id & '' AS id, mae_prov.numruc, mae_prov.nombre, Mid(con_percepcion!numreg,1,2)+mae_libros.codsun+Mid(con_percepcion!numreg,3,4) AS registro, 'Percepciones' AS libro, mae_documento.codsun, mae_documento.abrev, [con_percepcion]![numser]+'-'+[con_percepcion]![numdoc] AS numdoc2, con_percepcion.fchdoc ,con_percepcion.fchdoc as fchven, mae_moneda.simbolo, " _
''''                + vbCr + " con_tc.impven AS tipcam, con_percepcion.idmon, con_percepcion.imptotper AS imptotal,con_percepcion.impsal, " _
''''                + vbCr + " IIf(imptotal=0,0,IIf([con_percepcion].[idmon]=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, " _
''''                + vbCr + " IIf(imptotal=0,0,IIf([con_percepcion].[idmon]=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, " _
''''                + vbCr + " con_percepcion.glosa AS glosaope, 'H' as xOrigen" _
''''                + vbCr + " FROM (mae_moneda RIGHT JOIN (((con_percepcion LEFT JOIN mae_prov ON con_percepcion.idcli = mae_prov.id) LEFT JOIN mae_libros ON con_percepcion.idlib = mae_libros.id) LEFT JOIN mae_documento ON con_percepcion.tipdoc = mae_documento.id) ON mae_moneda.id = con_percepcion.idmon) LEFT JOIN con_tc ON con_percepcion.fchdoc = con_tc.fecha " _
''''                + vbCr + " WHERE (((con_percepcion.tipo)=1)) " & nSQLWhere
''''
''''            '--tabla visual que permitira dar un orden a la consulta
''''            nSQL = "SELECT tab.* FROM ( " & nSQL & " ) AS tab ORDER BY tab.nombre,tab.numdoc2 "
    
        Case 2 '--ventas
            '--Listado de facturacion, se incluye nc que esten den tesoreria origen ingreso
            nSQL = "SELECT vta_ventas.tipdoc,vta_ventas.id,IIf(vta_ventas!anulado=-1,' ',mae_cliente!numruc) AS numruc, IIf(vta_ventas!anulado=-1,'Anulado',mae_cliente!nombre) AS nombre, IIf(vta_ventas.numreg Is Null Or vta_ventas.numreg='',mae_libros.codsun,Left(vta_ventas.numreg,2) & mae_libros.codsun & Right(vta_ventas.numreg,4)) AS registro," _
                + vbCr + " 'Ventas' AS libro, mae_documento.codsun,mae_documento.abrev, vta_ventas!numser+'-'+vta_ventas!numdoc AS numdoc2, vta_ventas.fchdoc, vta_ventas.fchven, mae_moneda.simbolo, " _
                + vbCr + " IIf(vta_ventas.tc is null or vta_ventas.tc=0,con_tc.impven , vta_ventas.tc) AS tipcam,vta_ventas.idmon, " _
                + vbCr + " IIf(vta_ventas.numreg='000001',vta_ventas.imptotori,vta_ventas.imptotdoc) AS imptotal,vta_ventas.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(vta_ventas.idmon=1, imptotal ,IIf(tipcam Is Null,0, imptotal * tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(vta_ventas.idmon=2, imptotal ,IIf(tipcam Is Null,0, imptotal / tipcam))) AS imptotdol, " _
                + vbCr + " vta_ventas.glosa as glosaope, IIf(vta_ventas.tipdoc<>7,'D','H') as xOrigen" _
                + vbCr + " FROM ((((vta_ventas LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) " _
                + vbCr + "         LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id " _
                + vbCr + " WHERE ( (vta_ventas.anulado=0 AND IIf(vta_ventas.numreg='000001',vta_ventas.imptotori,vta_ventas.imptotdoc)<>0) ) " & nSQLWhere _
                + vbCr + " ORDER BY IIf(vta_ventas.anulado=-1,'Anulado',mae_cliente!nombre), vta_ventas!numser+'-'+vta_ventas!numdoc;"
        
'            nSQL = "SELECT vta_ventas.tipdoc,vta_ventas.id,IIf(vta_ventas!anulado=-1,' ',mae_cliente!numruc) AS numruc, IIf(vta_ventas!anulado=-1,'Anulado',mae_cliente!nombre) AS nombre, IIf(vta_ventas.numreg Is Null Or vta_ventas.numreg='',mae_libros.codsun,Left(vta_ventas.numreg,2) & mae_libros.codsun & Right(vta_ventas.numreg,4)) AS registro," _
                + vbCr + " 'Ventas' AS libro, mae_documento.codsun,mae_documento.abrev, vta_ventas!numser+'-'+vta_ventas!numdoc AS numdoc2, vta_ventas.fchdoc, vta_ventas.fchven, mae_moneda.simbolo, " _
                + vbCr + " IIf(vta_ventas.tc is null or vta_ventas.tc=0,con_tc.impven , vta_ventas.tc) AS tipcam,vta_ventas.idmon, " _
                + vbCr + " IIf(vta_ventas.numreg='000001',vta_ventas.imptotori,vta_ventas.imptotdoc) AS imptotal,vta_ventas.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(vta_ventas.idmon=1, imptotal ,IIf(tipcam Is Null,0, imptotal * tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(vta_ventas.idmon=2, imptotal ,IIf(tipcam Is Null,0, imptotal / tipcam))) AS imptotdol, " _
                + vbCr + " vta_ventas.glosa as glosaope, IIf(vta_ventas.tipdoc<>7,'D','H') as xOrigen" _
                + vbCr + " FROM ( ((((vta_ventas LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) " _
                + vbCr + "         LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id " _
                + vbCr + "       ) LEFT JOIN " _
                + vbCr + " ( SELECT tes_cajaorigendet.iddoc FROM tes_cajaorigendet INNER JOIN vta_ventas ON tes_cajaorigendet.iddoc = vta_ventas.id WHERE (((tes_cajaorigendet.idmod)=1) AND ((vta_ventas.tipdoc)=7)) " _
                + vbCr + "   UNION " _
                + vbCr + "   SELECT tes_cajadestinodet.iddoc FROM tes_cajadestinodet INNER JOIN vta_ventas ON tes_cajadestinodet.iddoc = vta_ventas.id WHERE (((tes_cajadestinodet.idmod)=1) AND ((vta_ventas.tipdoc)=7)) " _
                + vbCr + "  ) as tes " _
                + vbCr + " ON vta_ventas.id = tes.iddoc " _
                + vbCr + " WHERE ( (vta_ventas.tipdoc<>7 AND vta_ventas.anulado=0 AND IIf(vta_ventas.numreg='000001',vta_ventas.imptotori,vta_ventas.imptotdoc)<>0) OR " _
                + vbCr + "        (vta_ventas.tipdoc=7 AND vta_ventas.anulado=0 AND IIf(vta_ventas.numreg='000001',vta_ventas.imptotori,vta_ventas.imptotdoc)<>0 AND vta_ventas.iddocref=0) OR " _
                + vbCr + "        (tes.iddoc is not null) ) " & nSQLWhere _
                + vbCr + " ORDER BY IIf(vta_ventas.anulado=-1,'Anulado',mae_cliente!nombre), vta_ventas!numser+'-'+vta_ventas!numdoc;"
        
        Case 4 '--Percepciones
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "con_percepcion")
            nSQLWhere = Replace(nSQLWhere, "idpro", "idcli")
            
            nSQL = "SELECT con_percepcion.id AS id, con_percepcion.tipdoc, mae_prov.numruc, mae_prov.nombre AS nombre, IIf([con_percepcion].[numreg] Is Null Or [con_percepcion].[numreg]='',[mae_libros].[codsun],Left([con_percepcion].[numreg],2) & [mae_libros].[codsun] & Right([con_percepcion].[numreg],4)) AS registro, mae_documento.abrev, " _
                + vbCr + " [con_percepcion].[numser]+'-'+[con_percepcion].[numdoc] AS numdoc2, con_percepcion.fchdoc,Null as fchven, mae_moneda.simbolo, " _
                + vbCr + " IIf([con_percepcion].[tc] Is Null Or [con_percepcion].[tc]=0,[con_tc].[impven],[con_percepcion].[tc]) AS tipcam, con_percepcion.idmon, con_percepcion.imptotper AS imptotal, con_percepcion.impsal, " _
                + vbCr + " IIf([imptotal]=0,0,IIf([con_percepcion].[idmon]=1,[imptotal],IIf([tipcam] Is Null,0,[imptotal]*[tipcam]))) AS imptotsol, " _
                + vbCr + " IIf([imptotal]=0,0,IIf([con_percepcion].[idmon]=2,[imptotal],IIf([tipcam] Is Null,0,[imptotal]/[tipcam]))) AS imptotdol, " _
                + vbCr + " con_percepcion.glosa as glosaope, '' AS docref, 'Percepciones' AS libro,'H' as xOrigen  " _
                + vbCr + " FROM ((((con_percepcion LEFT JOIN mae_documento ON con_percepcion.tipdoc = mae_documento.id) LEFT JOIN mae_libros ON con_percepcion.idlib = mae_libros.id) LEFT JOIN con_tc ON con_percepcion.fchdoc = con_tc.fecha) LEFT JOIN mae_prov ON con_percepcion.idcli = mae_prov.id) LEFT JOIN mae_moneda ON con_percepcion.idmon = mae_moneda.id " _
                + vbCr + " Where (((con_percepcion.imptotper) <> 0)) " & nSQLWhere _
                + vbCr + " ORDER BY IIf([con_percepcion].[numreg] Is Null Or [con_percepcion].[numreg]='',[mae_libros].[codsun],Left([con_percepcion].[numreg],2) & [mae_libros].[codsun] & Right([con_percepcion].[numreg],4)), con_percepcion.fchdoc  "

        
        Case 40 '--honorarios
            '--reemplazar filtro
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "com_honorarios")
            nSQLWhere = Replace(nSQLWhere, "idcli", "idpro")
            
            nSQL = "SELECT com_honorarios.tipdoc, com_honorarios.id,mae_prov.numruc, mae_prov!nombre AS nombre, IIf(com_honorarios.numreg Is Null Or com_honorarios.numreg='',mae_libros.codsun ,Left(com_honorarios.numreg,2) & mae_libros.codsun & Right(com_honorarios.numreg,4)) AS registro, 'Honorario' AS libro, mae_documento.codsun,mae_documento.abrev, com_honorarios!numser+'-'+com_honorarios!numdoc AS numdoc2, com_honorarios.fchdoc, com_honorarios.fchven, mae_moneda.simbolo, " _
                + vbCr + " iif(com_honorarios.tc is null or com_honorarios.tc=0,con_tc.impven , com_honorarios.tc) AS tipcam,com_honorarios.idmon, " _
                + vbCr + " IIf(com_honorarios.numreg='000001',com_honorarios.imptotori,com_honorarios.imptot) AS imptotal, com_honorarios.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf([com_honorarios].[idmon]=1, imptotal ,IIf(tipcam Is Null,0, imptotal * tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf([com_honorarios].[idmon]=2, imptotal ,IIf(tipcam Is Null,0, imptotal / tipcam))) AS imptotdol, " _
                + vbCr + " com_honorarios.glosa as glosaope, 'H' as xOrigen" _
                + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN (mae_documento RIGHT JOIN (com_honorarios LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) ON mae_documento.id = com_honorarios.tipdoc) ON mae_prov.id = com_honorarios.idpro) ON mae_moneda.id = com_honorarios.idmon) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha " _
                + vbCr + " WHERE  (IIf([com_honorarios].[numreg]='000001',[com_honorarios].[imptotori],[com_honorarios].[imptot])<>0) and ( com_honorarios.tipdoc <> 7) " & nSQLWhere _
                + vbCr + " ORDER BY mae_prov!nombre, com_honorarios.fchdoc;"
    
        Case 9 '--Remuneraciones
            '--reemplazar filtro
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "pla_boleta")
            nSQLWhere = Replace(nSQLWhere, "idcli", "idemp")
                        
            nSQL = "SELECT pla_boleta.iddoc as tipdoc, pla_boleta.id,pla_empleados.numdoc as numruc, pla_empleados.nombre AS nombre, IIf(pla_boleta.numreg Is Null Or pla_boleta.numreg='',mae_libros.codsun ,Left(pla_boleta.numreg,2) & mae_libros.codsun & Right(pla_boleta.numreg,4)) AS registro, 'Remuneraciones' AS libro, mae_documento.codsun,mae_documento.abrev, pla_boleta.numser+'-'+pla_boleta.numdoc AS numdoc2, pla_boleta.fchdoc, null as fchven, mae_moneda.simbolo, " _
                + vbCr + " con_tc.impven AS tipcam,pla_boleta.idmon, " _
                + vbCr + " IIf(pla_boleta.numreg='000001',pla_boleta.imptot,pla_boleta.imptot) AS imptotal, pla_boleta.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(pla_boleta.idmon=1, imptotal ,IIf(tipcam Is Null,0, imptotal * tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(pla_boleta.idmon=2, imptotal ,IIf(tipcam Is Null,0, imptotal / tipcam))) AS imptotdol, " _
                + vbCr + " pla_boleta.glosa as glosaope,'H' as xOrigen" _
                + vbCr + " FROM (mae_moneda RIGHT JOIN (pla_empleados RIGHT JOIN (mae_documento RIGHT JOIN (pla_boleta LEFT JOIN mae_libros ON pla_boleta.idlib = mae_libros.id) ON mae_documento.id = pla_boleta.iddoc) ON pla_empleados.id = pla_boleta.idemp) ON mae_moneda.id = pla_boleta.idmon) LEFT JOIN con_tc ON pla_boleta.fchdoc = con_tc.fecha " _
                + vbCr + " WHERE  (IIf(pla_boleta.numreg='000001',pla_boleta.imptot,pla_boleta.imptot)<>0) and ( pla_boleta.iddoc <> 7) " & nSQLWhere _
                + vbCr + " ORDER BY pla_empleados.nombre, pla_boleta.fchdoc;"
    
        Case 37 '--Letras
            '--reemplazar filtro
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "let_letra")
            nSQLWhere = Replace(nSQLWhere, "idcli", "idclipro")
            
            nSQL = "SELECT let_letradet.corr AS id, mae_cliente.numruc, mae_cliente.nombre, Left([let_letra].[numreg],2) & [mae_libros].[codsun] & Right([let_letra].[numreg],4) AS registro, 'Letras' AS libro, mae_documento.abrev, [let_letra].[ano] & ' ' & [let_letradet].[numdoc] & ' ' & [let_letradet].[numser] AS numdoc2, " _
                + vbCr + " let_letradet.fchemi AS fchdoc, let_letradet.fchven, mae_moneda.simbolo, IIf([let_letra].[tc]=0,[con_tc].[impven],[let_letra].[tc]) AS tipcam, let_letra.idmon, let_letradet.implet AS imptotal, let_letradet.impsal, " _
                + vbCr + " IIf(let_letra.idmon=1,let_letradet.implet,let_letradet.implet*tipcam) AS imptotsol, " _
                + vbCr + " IIf(let_letra.idmon=2,let_letradet.implet,IIf(tipcam=0,0,let_letradet.implet/tipcam)) AS imptotdol, " _
                + vbCr + " let_letra.glosa AS glosaope, 'D' as xOrigen" _
                + vbCr + " FROM mae_moneda RIGHT JOIN (((((mae_cliente RIGHT JOIN let_letra ON mae_cliente.id = let_letra.idclipro) LEFT JOIN mae_documento ON let_letra.tipdoc = mae_documento.id) LEFT JOIN mae_libros ON let_letra.idlib = mae_libros.id) LEFT JOIN con_tc ON let_letra.fchemi = con_tc.fecha) INNER JOIN let_letradet ON let_letra.id = let_letradet.idlet) ON mae_moneda.id = let_letra.idmon " _
                + vbCr + " WHERE (((let_letradet.fchemi)<=CDate('" & TxtFchFin.Valor & "'))) " & nSQLWhere _
                + vbCr + " ORDER BY mae_cliente.nombre, [let_letra].[ano] & ' ' & [let_letradet].[numdoc] & ' ' & [let_letradet].[numser];"
        
        Case 41 '--Lgd
            '--reemplazar filtro
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "vta_gastodebito")
           
            nSQL = "SELECT vta_gastodebito.id,IIf([vta_gastodebito]![anulado]=-1,' ',[mae_cliente]![numruc]) AS numruc, IIf([vta_gastodebito]![anulado]=-1,'Anulado',[mae_cliente]![nombre]) AS nombre, IIf([vta_gastodebito].[numreg] Is Null Or [vta_gastodebito].[numreg]='',[mae_libros].[codsun],Left([vta_gastodebito].[numreg],2) & [mae_libros].[codsun] & Right([vta_gastodebito].[numreg],4)) AS registro, 'Lgd' AS libro, mae_documento.codsun,mae_documento.abrev, [vta_gastodebito]![numser]+'-'+[vta_gastodebito]![numdoc] AS numdoc2, " _
                + vbCr + " vta_gastodebito.fchemi as fchdoc, vta_gastodebito.fchven, mae_moneda.simbolo, " _
                + vbCr + " iif(vta_gastodebito.tc is null or vta_gastodebito.tc=0,con_tc.impven , vta_gastodebito.tc) AS tipcam, " _
                + vbCr + " vta_gastodebito.idmon,vta_gastodebito.imptot AS imptotal,vta_gastodebito.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf([vta_gastodebito].[idmon]=1, imptotal ,IIf(tipcam Is Null,0, imptotal * tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf([vta_gastodebito].[idmon]=2, imptotal ,IIf(tipcam Is Null,0, imptotal / tipcam))) AS imptotdol, " _
                + vbCr + " vta_gastodebito.glosa as glosaope, IIf(vta_gastodebito.tipdoc<>126,'D','H' ) as xOrigen " _
                + vbCr + " FROM ((((vta_gastodebito LEFT JOIN mae_cliente ON vta_gastodebito.idcli = mae_cliente.id) LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha) LEFT JOIN mae_libros ON vta_gastodebito.idlib = mae_libros.id) LEFT JOIN mae_moneda ON vta_gastodebito.idmon = mae_moneda.id " _
                + vbCr + " WHERE ( (vta_gastodebito.tipdoc <>126 and  vta_gastodebito.anulado =0) ) " & nSQLWhere _
                + vbCr + " ORDER BY IIf([vta_gastodebito]![anulado]=-1,'Anulado',[mae_cliente]![nombre]), [vta_gastodebito]![numser]+'-'+[vta_gastodebito]![numdoc];"
        
        Case 42 '--Planilla Letras
            '--reemplazar filtro
            nSQLWhere = Replace(nSQLWhere, "ventas.idcli", "mae_banco.idban")
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "let_planilla")
            If nSQLWhere <> "" Then nSQLWhere = "Where" & Mid(nSQLWhere, 5)
            
            nSQL = "SELECT let_planilla.id, mae_bancos.numruc, mae_bancos.descripcion AS nombre, Left([let_planilla].[numreg],2) & [mae_libros].[codsun] & Right([let_planilla].[numreg],4) AS registro, 'Planilla letra' AS libro, mae_documento.abrev, let_planilla.numdoc AS numdoc2, " _
            + vbCr + " let_planilla.fchemi AS fchdoc, '' AS fchven, mae_moneda.simbolo, IIf([let_planilla].[anulado]=-1,0,IIf([let_planilla].[tc]=0,[con_tc].[impven],[let_planilla].[tc])) AS tipcam, let_planilla.idmon, let_planilla.imptot AS imptotal,  let_planilla.impsal," _
            + vbCr + " IIf(let_planilla.idmon=1,let_planilla.imptot,let_planilla.imptot*tipcam) AS imptotsol, " _
            + vbCr + " IIf(let_planilla.idmon=2,let_planilla.imptot,IIf(tipcam=0,0,let_planilla.imptot/tipcam)) AS imptotdol , " _
            + vbCr + " let_planilla.glosa AS glosaope, 'D' as xOrigen " _
            + vbCr + " FROM mae_documento RIGHT JOIN (mae_bancos RIGHT JOIN ((((let_planilla LEFT JOIN mae_moneda ON let_planilla.idmon = mae_moneda.id) LEFT JOIN mae_banconumcta ON let_planilla.idbcocta = mae_banconumcta.id) LEFT JOIN mae_libros ON let_planilla.idlib = mae_libros.id) LEFT JOIN con_tc ON let_planilla.fchemi = con_tc.fecha) ON mae_bancos.id = mae_banconumcta.idban) ON mae_documento.id = let_planilla.tipdoc " _
            + vbCr + nSQLWhere _
            + vbCr + " ORDER BY mae_bancos.descripcion, let_planilla.numdoc;"
        
        Case 999 '--Reembolsables
            '--reemplazar filtro
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "com_reembolsables")
            nSQLWhere = Replace(nSQLWhere, "idcli", "idpro")
            nSQLWhere = Replace(nSQLWhere, "fchreg", "fchdoc")
         
            nSQL = "SELECT  com_reembolsables.id,mae_prov.numruc, mae_prov!nombre AS nombre, IIf(com_reembolsables.numreg Is Null Or com_reembolsables.numreg='',mae_libros.codsun ,Left(com_reembolsables.numreg,2) & mae_libros.codsun & Right(com_reembolsables.numreg,4)) AS registro, 'Reembolsables' AS libro, mae_documento.codsun,mae_documento.abrev, iif(com_reembolsables!numser is null or com_reembolsables!numser ='','',com_reembolsables!numser  +'-' ) + com_reembolsables!numdoc AS numdoc2, com_reembolsables.fchdoc, com_reembolsables.fchven, mae_moneda.simbolo, " _
                + vbCr + " iif(com_reembolsables.tc is null or com_reembolsables.tc=0,con_tc.impven , com_reembolsables.tc) AS tipcam, " _
                + vbCr + " com_reembolsables.idmon,com_reembolsables.imptot AS imptotal,com_reembolsables.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf([com_reembolsables].[idmon]=1, imptotal ,IIf(tipcam Is Null,0, imptotal * tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf([com_reembolsables].[idmon]=2, imptotal ,IIf(tipcam Is Null,0, imptotal / tipcam))) AS imptotdol, " _
                + vbCr + " com_reembolsables.glosa as glosaope, IIf(com_reembolsables.tipdoc<>7,'H','D' ) as xOrigen " _
                + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN (mae_documento RIGHT JOIN (com_reembolsables LEFT JOIN mae_libros ON com_reembolsables.idlib = mae_libros.id) ON mae_documento.id = com_reembolsables.tipdoc) ON mae_prov.id = com_reembolsables.idpro) ON mae_moneda.id = com_reembolsables.idmon) LEFT JOIN con_tc ON com_reembolsables.fchdoc = con_tc.fecha " _
                + vbCr + " WHERE ( com_reembolsables.tipdoc <> 7) " & nSQLWhere _
                + vbCr + " ORDER BY mae_prov!nombre, com_reembolsables.fchdoc "
            
            '--tabla vistual que permitira dar un orden a la consulta
            nSQL = "SELECT tab.* FROM ( " & nSQL & " ) AS tab ORDER BY tab.nombre,tab.numdoc2 "
            
        Case Else
            Exit Sub

    End Select
    
    
    '--indicar el campo a mostrar segun la moneda seleccionada
    If NulosN(TxtIdMon.Text) = 1 Then
        nCampoMuestra = "imptotsol"
    ElseIf NulosN(TxtIdMon.Text) = 2 Then
        nCampoMuestra = "imptotdol"
    Else
        fraBarra.Visible = False
        MsgBox "Por el momento no se puede expresar en " & LblMoneda.Caption, vbInformation, xTitulo
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    '--ejecutar la conulta
    RST_Busq rst, nSQL, xCon
    If rst.State = 0 Then GoTo Salir
    
    '--filtrar lo que se va mostrar
    If chk_descuadrado.Value = 0 Then
        '--obs. si selecciona la opcion todos no hace el fintro
        If OptPen.Value = True Then rst.Filter = "impsal > 0" ' FILTRAMOS LOS PENDIENTE
        If OptCan.Value = True Then rst.Filter = "impsal <= 0" ' FILTRAMOS LOS CANCELADOS
    End If
    
    If rst.RecordCount = 0 Then
        MsgBox "No hay documentos del " & Fg4.TextMatrix(Fg4.Row, 1) & " seleccionado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fraBarra.Visible = False
        Set rst = Nothing
        Exit Sub
    End If
    
    '--aplicando orden
    
    '--si muestra todos los clientes,proveedores
    If OptSel1.Value = True Then
        If Opt_Orden(0).Value = True Then '--numero doc
            rst.Sort = "nombre,numdoc2"
        ElseIf Opt_Orden(1).Value = True Then '--registro
            rst.Sort = "nombre,registro"
        ElseIf Opt_Orden(2).Value = True Then '--fecha doc
            rst.Sort = "nombre,fchdoc,numdoc2"
        Else
            rst.Sort = "nombre,numdoc2,fchdoc"
        End If
    Else
        If Opt_Orden(0).Value = True Then '--numero doc
            rst.Sort = "numdoc2"
        ElseIf Opt_Orden(1).Value = True Then '--registro
            rst.Sort = "registro"
        ElseIf Opt_Orden(2).Value = True Then '--fecha doc
            rst.Sort = "fchdoc,numdoc2"
        Else
            rst.Sort = "numdoc2,fchdoc"
        End If
    End If
    '-------------------------------------
    
    ProgressBar1.Max = rst.RecordCount
    
    Dim xSaldoDoc As Double
    Dim xFilaIni&
    Dim xFilaIniGrupo As Long '--almacena la fila de inicio de grupo proveedor/ciente/prestador servicio
    Dim xColor&
    
    Me.MousePointer = vbHourglass
     
    xColor = 0
    
    
    If rst.RecordCount <> 0 Then
        DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then GoTo Salir:

        rst.MoveFirst
        xSaldoDoc = 0
        xNomPro = NulosC(rst("nombre"))
        xFila = Fg1.FixedRows
        
        '------------------------------------------------------------------------
        '--colocar datos del grupo(cliente, proveedor o prestador de servicio
        '--detalle
        Fg1.Rows = Fg1.Rows + 1

        GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 12, "Nº R.U.C. : " & RellenarBlancos(NulosC(rst("numruc")), 12, 1) & "  " & xNomPro, flexAlignLeftCenter, True, , , , True
        xFilaIniGrupo = Fg1.Rows - 1
        
        '--resumen
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(rst("numruc"))
        Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosC(rst("nombre"))
        '------------------------------------------------------------------------
        
        xFilaIni = xFila
        
        'dar formato a la fila
        With Fg1
            .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H800000
            .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
            .FillStyle = flexFillRepeat
            .CellFontBold = True
        End With
    
        TotDebe = 0
        TotHaber = 0
        
        Cambio = False
        
        Dim mRowIni As Integer
        '--considerar los id de libros
        '         1=Compras
        '         2=Ventas
        '         5=Igv Retenciones;
        '         6=Bancos;
        '         8=Canjes de Facturas
        '         39=Rendición de Cuentas
        '         40=Registro de Honorarios Profesionales

'-----------------------------------------------------------
    nSQL = ""
    If IdCliPro <> 0 Then nSQL = " and con_diario.ridper = " & IdCliPro & " "
        
    nSQL = "SELECT con_diario.rregistro, Format(con_diario.idmes,'00') & mae_libros.codsun & Format(con_diario.numasi,'0000') AS registro, mae_libros.descripcion AS libro, IIf(con_diario.ridtipper2=5,mae_bancos.abrev, " _
        + vbCr + " IIf(con_diario.ridtipper2=2,mae_cliente.nombre,IIf(con_diario.ridtipper2=1,mae_prov.nombre,''))) AS razonsocial, tes_documentos.abrev, con_diario.rnumerodoc2 AS numdoc, con_diario.rfchope2 AS fchemi, mae_moneda.simbolo, " _
        + vbCr + " IIf(con_diario.aplicatc=0,con_tc.impven,con_diario.tc) AS tipcam, " _
        + vbCr + " IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS ximptotal, " _
        + vbCr + " IIf(con_diario.idmon=1,con_diario.impdebsol,con_diario.impdebdol) AS impdeb, " _
        + vbCr + " IIf(con_diario.idmon=1,con_diario.imphabsol,con_diario.imphabdol) AS imphab, " _
        + vbCr + " IIf(impdeb=0,imphab,impdeb) as imptotal, " _
        + vbCr + " IIf(con_diario.idmon=1,imptotal,imptotal*tipcam)   AS imptotsol," _
        + vbCr + " IIf(con_diario.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam))  AS imptotdol, " _
        + vbCr + " con_diario.ridper, con_diario.rnumerodoc AS numdoc2, con_diario.rglosaope, con_diario.iddoc, con_diario.idmon " _
        + vbCr + " FROM ((((((con_diario LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN tes_documentos ON con_diario.rtipdoc2 = tes_documentos.id) " _
        + vbCr + " LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN mae_bancos ON con_diario.ridper2 = mae_bancos.id) LEFT JOIN mae_cliente ON con_diario.ridper2 = mae_cliente.id) LEFT JOIN mae_prov ON con_diario.ridper2 = mae_prov.id " _
        + vbCr + " WHERE (((con_diario.idlib) In (" & xIdLibroRef & ")) AND ((con_diario.ridlib)=" & xIdLibro & ")) " & nSQL & nSQLAjuste
        
        If xIdLibro = 1 Then
        
        '--Verificar si hay documentos de NC que fueron registrados en Tesoreria Ingresos - Egresos
        nSQLDocNCBancos = BuscarNCBancos()
        If nSQLDocNCBancos <> "" Then
            nSQLDocNCBancos = " and com_compras.id not in (" & nSQLDocNCBancos & ")"
        End If
        
        '--Cancelacion de compras con nota de credito excepto nc que se registran en tesoreria
        nSQL = nSQL _
            + vbCr + " UNION " _
            + vbCr + " SELECT Left(com_compras_1.numreg,2) & mae_libros_1.codsun & Right(com_compras_1.numreg,4) AS rregistro, Mid(com_compras!numreg,1,2) & mae_libros!codsun & Mid(com_compras.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_prov.nombre AS razonsocial, mae_documento.abrev, " _
            + vbCr + " com_compras.numser & '-' & com_compras.numdoc AS numdoc, com_compras.fchdoc AS fchemi, mae_moneda.simbolo, IIf(com_compras_1.tc=0,con_tc.impven,com_compras_1.tc) AS tipcam, " _
            + vbCr + " com_compras.imptot AS ximptotal,ximptotal as impdeb,0 as imphab,ximptotal as imptotal, " _
            + vbCr + " IIf(com_compras.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, " _
            + vbCr + " IIf(com_compras.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) AS imptotdol, " _
            + vbCr + " com_compras.idpro AS ridper, com_compras_1.numser & '-' & com_compras_1.numdoc AS numdoc2, com_compras.glosa AS glosaope, com_compras_1.id AS iddoc, com_compras.idmon " _
            + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((((com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) INNER JOIN com_compras AS com_compras_1 ON com_compras.iddocref = com_compras_1.id)  " _
            + vbCr + " LEFT JOIN mae_libros AS mae_libros_1 ON com_compras_1.idlib = mae_libros_1.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) LEFT JOIN mae_prov ON com_compras.idpro = mae_prov.id " _
            + vbCr + " WHERE (com_compras.iddocref<>0 ) " & IIf(IdCliPro <> 0, " and com_compras.idpro=" & IdCliPro, "") & nSQLDocNCBancos

        '--Cancelacion de notas de credito con compras excepto nc que se registran en tesoreria
        nSQLDocNCBancos = Replace(nSQLDocNCBancos, "com_compras", "com_compras_1")

        nSQL = nSQL _
            + vbCr + " UNION " _
            + vbCr + " SELECT Left(com_compras_1.numreg,2) & mae_libros_1.codsun & Right(com_compras_1.numreg,4) AS rregistro, Mid(com_compras!numreg,1,2) & mae_libros!codsun & Mid(com_compras.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_prov.nombre AS razonsocial, mae_documento.abrev, " _
            + vbCr + " com_compras.numser & '-' & com_compras.numdoc AS numdoc, com_compras.fchdoc AS fchemi, mae_moneda.simbolo, IIf(com_compras_1.tc=0,con_tc.impven,com_compras_1.tc) AS tipcam,  " _
            + vbCr + " com_compras_1.imptot AS ximptotal, 0 as impdeb,ximptotal as imphab,ximptotal as imptotal, " _
            + vbCr + " IIf(com_compras.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, " _
            + vbCr + " IIf(com_compras.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) AS imptotdol, " _
            + vbCr + " com_compras.idpro AS ridper, com_compras_1.numser & '-' & com_compras_1.numdoc AS numdoc2, com_compras.glosa AS glosaope, com_compras_1.id AS iddoc, com_compras.idmon " _
            + vbCr + " FROM (com_compras AS com_compras_1 LEFT JOIN mae_libros AS mae_libros_1 ON com_compras_1.idlib = mae_libros_1.id) INNER JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (((com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) " _
            + vbCr + " LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) LEFT JOIN mae_prov ON com_compras.idpro = mae_prov.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON com_compras_1.iddocref = com_compras.id " _
            + vbCr + " WHERE (com_compras_1.iddocref<>0 ) " & IIf(IdCliPro <> 0, " and com_compras_1.idpro=" & IdCliPro, "") & nSQLDocNCBancos
        '--------------------------------------------------

        
        
''            nSQLWhere = Replace(nSQLWhere, "con_percepcion", "com_compras")
''            If IdCliPro <> 0 Then nSQLWhere = " and com_compras.idpro = " & IdCliPro & " "
''           nSQL = nSQL + vbCr + " Union All " _
''                + vbCr + "SELECT Left(com_compras_1.numreg,2) & mae_libros_1.codsun & Right(com_compras_1.numreg,4) AS rregistro, Mid(com_compras.numreg,1,2) & mae_libros.codsun & Mid(com_compras.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_prov.nombre, mae_documento.abrev, " _
''                + vbcr + " com_compras.numser & '-' & com_compras.numdoc AS numdoc, com_compras.fchdoc AS fchemi, mae_moneda.simbolo, IIf(com_compras_1.tc=0,con_tc.impven,com_compras_1.tc) AS tipcam, " _
''                + vbCr + " com_compras.imptot AS ximptotal, ximptotal as impdeb,0 as imphab,ximptotal as imptotal, " _
''                + vbCr + " IIf(com_compras.idmon=1,com_compras.imptot,com_compras.imptot*tipcam) AS imptotsol, " _
''                + vbCr + " IIf(com_compras.idmon=2,com_compras.imptot,iif(tipcam=0,0,com_compras.imptot/tipcam)) AS imptotdol, " _
''                + vbCr + " com_compras.idpro AS ridper,  com_compras_1.numser & '-' & com_compras_1.numdoc AS numdoc2, com_compras.glosa AS glosaope, com_compras_1.id AS iddoc, com_compras_1.id AS idmon  " _
''                + vbCr + " FROM ( ((mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((((com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) INNER JOIN com_compras AS com_compras_1 ON com_compras.iddocref = com_compras_1.id) LEFT JOIN mae_libros AS mae_libros_1 ON com_compras_1.idlib = mae_libros_1.id) " _
''                + vbCr + "       ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) LEFT JOIN (con_planctas RIGHT JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (com_compras.idmon = mae_documentocta.idmon) AND (com_compras.tipdoc = mae_documentocta.iddoc)) LEFT JOIN mae_prov ON com_compras.idpro = mae_prov.id " _
''                + vbCr + "        ) LEFT JOIN " _
''                + vbCr + " ( SELECT tes_cajaorigendet.iddoc FROM tes_cajaorigendet INNER JOIN com_compras ON tes_cajaorigendet.iddoc = com_compras.id WHERE (((tes_cajaorigendet.idmod)=1) AND ((com_compras.tipdoc)=7)) " _
''                + vbCr + "   ) as tes ON com_compras.id=tes.iddoc  " _
''                + vbCr + " WHERE com_compras.iddocref Is Not Null And com_compras.iddocref<>0 and mae_documentocta.tipope=0 and tes.iddoc is null " & nSQLWhere
        
        ElseIf xIdLibro = 2 Then
        
        '--Verificar si hay documentos de NC que fueron registrados en Tesoreria Ingresos - Egresos
        nSQLDocNCBancos = BuscarNCBancos()
        If nSQLDocNCBancos <> "" Then
            nSQLDocNCBancos = " and vta_ventas.id not in (" & nSQLDocNCBancos & ")"
        End If
        
        '--Cancelacion de ventas con nota de credito excepto nc que se registran en tesoreria
        nSQL = nSQL _
            + vbCr + " UNION " _
            + vbCr + " SELECT Left(vta_ventas_1.numreg,2) & mae_libros_1.codsun & Right(vta_ventas_1.numreg,4) AS rregistro, Mid(vta_ventas!numreg,1,2) & mae_libros!codsun & Mid(vta_ventas.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_prov.nombre AS razonsocial, mae_documento.abrev, " _
            + vbCr + " vta_ventas.numser & '-' & vta_ventas.numdoc AS numdoc, vta_ventas.fchdoc AS fchemi, mae_moneda.simbolo, IIf(vta_ventas_1.tc=0,con_tc.impven,vta_ventas_1.tc) AS tipcam, " _
            + vbCr + " vta_ventas.imptotdoc AS ximptotal,0 as impdeb,ximptotal as imphab,ximptotal as imptotal, " _
            + vbCr + " IIf(vta_ventas.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, " _
            + vbCr + " IIf(vta_ventas.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) AS imptotdol, " _
            + vbCr + " vta_ventas.idcli AS ridper, vta_ventas_1.numser & '-' & vta_ventas_1.numdoc AS numdoc2, vta_ventas.glosa AS glosaope, vta_ventas_1.id AS iddoc, vta_ventas.idmon " _
            + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((((vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) INNER JOIN vta_ventas AS vta_ventas_1 ON vta_ventas.iddocref = vta_ventas_1.id)  " _
            + vbCr + " LEFT JOIN mae_libros AS mae_libros_1 ON vta_ventas_1.idlib = mae_libros_1.id) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) LEFT JOIN mae_prov ON vta_ventas.idcli = mae_prov.id " _
            + vbCr + " WHERE (vta_ventas.iddocref<>0 ) " & IIf(IdCliPro <> 0, " and vta_ventas.idcli=" & IdCliPro, "") & nSQLDocNCBancos

        '--Cancelacion de notas de credito con ventas excepto nc que se registran en tesoreria
        nSQLDocNCBancos = Replace(nSQLDocNCBancos, "vta_ventas", "vta_ventas_1")

        nSQL = nSQL _
            + vbCr + " UNION " _
            + vbCr + " SELECT Left(vta_ventas_1.numreg,2) & mae_libros_1.codsun & Right(vta_ventas_1.numreg,4) AS rregistro, Mid(vta_ventas!numreg,1,2) & mae_libros!codsun & Mid(vta_ventas.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_prov.nombre AS razonsocial, mae_documento.abrev, " _
            + vbCr + " vta_ventas.numser & '-' & vta_ventas.numdoc AS numdoc, vta_ventas.fchdoc AS fchemi, mae_moneda.simbolo, IIf(vta_ventas_1.tc=0,con_tc.impven,vta_ventas_1.tc) AS tipcam,  " _
            + vbCr + " vta_ventas_1.imptotdoc AS ximptotal, ximptotal as impdeb,0 as imphab,ximptotal as imptotal, " _
            + vbCr + " IIf(vta_ventas.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, " _
            + vbCr + " IIf(vta_ventas.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) AS imptotdol, " _
            + vbCr + " vta_ventas.idcli AS ridper, vta_ventas_1.numser & '-' & vta_ventas_1.numdoc AS numdoc2, vta_ventas.glosa AS glosaope, vta_ventas_1.id AS iddoc, vta_ventas.idmon " _
            + vbCr + " FROM (vta_ventas AS vta_ventas_1 LEFT JOIN mae_libros AS mae_libros_1 ON vta_ventas_1.idlib = mae_libros_1.id) INNER JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (((vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) " _
            + vbCr + " LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_prov ON vta_ventas.idcli = mae_prov.id) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON vta_ventas_1.iddocref = vta_ventas.id " _
            + vbCr + " WHERE (vta_ventas_1.iddocref<>0 ) " & IIf(IdCliPro <> 0, " and vta_ventas_1.idcli=" & IdCliPro, "") & nSQLDocNCBancos
        '--------------------------------------------------
        
''            '--unido a referencias de nota de credito
''            If IdCliPro <> 0 Then nSQLWhere = " and vta_ventas.idcli = " & IdCliPro & " "
''            '--el tipo de cambio de la NC se obtendra del documento de referencia, en caso de no tener ingresado manualmente
''            nSQL = nSQL + vbCr + " Union All " _
''                + vbCr + "SELECT Left(vta_ventas_1.numreg,2) & mae_libros_1.codsun & Right(vta_ventas_1.numreg,4) AS rregistro, Mid(vta_ventas.numreg,1,2) & mae_libros.codsun & Mid(vta_ventas.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_cliente.nombre, mae_documento.abrev, vta_ventas.numser & '-' & vta_ventas.numdoc AS numdoc, vta_ventas.fchdoc AS fchemi, mae_moneda.simbolo, IIf(vta_ventas_1.tc=0,con_tc.impven,vta_ventas_1.tc) AS tipcam, " _
''                + vbCr + " vta_ventas.imptotdoc AS ximptotal, 0 as impdeb,ximptotal as imphab,ximptotal as imptotal, " _
''                + vbCr + " IIf(vta_ventas.idmon=1,vta_ventas.imptotdoc,vta_ventas.imptotdoc*tipcam) AS imptotsol, " _
''                + vbCr + " IIf(vta_ventas.idmon=2,vta_ventas.imptotdoc,iif(tipcam=0,0,vta_ventas.imptotdoc/tipcam) ) AS imptotdol, " _
''                + vbCr + " vta_ventas.idcli AS ridper,  vta_ventas_1.numser & '-' & vta_ventas_1.numdoc AS numdoc2, vta_ventas.glosa AS glosaope, vta_ventas_1.id AS iddoc,vta_ventas_1.id AS idmon " _
''                + vbCr + " FROM (  ((((mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) INNER JOIN vta_ventas AS vta_ventas_1 ON vta_ventas.iddocref = vta_ventas_1.id) " _
''                + vbCr + "         LEFT JOIN mae_libros AS mae_libros_1 ON vta_ventas_1.idlib = mae_libros_1.id) LEFT JOIN (con_planctas RIGHT JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (vta_ventas.tipdoc = mae_documentocta.iddoc) AND (vta_ventas.idmon = mae_documentocta.idmon)) LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id " _
''                + vbCr + "       ) LEFT JOIN " _
''                + vbCr + " ( SELECT tes_cajaorigendet.iddoc FROM tes_cajaorigendet INNER JOIN vta_ventas ON tes_cajaorigendet.iddoc = vta_ventas.id WHERE (((tes_cajaorigendet.idmod)=2) AND ((vta_ventas.tipdoc)=7)) GROUP BY tes_cajaorigendet.iddoc " _
''                + vbCr + "   ) as tes ON vta_ventas.id = tes.iddoc " _
''                + vbCr + " WHERE vta_ventas.anulado=0 and vta_ventas.tipdoc=7 and vta_ventas.iddocref <> 0 and mae_documentocta.tipope =-1 and  tes.iddoc is null " & nSQLWhere
        
        End If
        
    RST_Busq Rstabo, nSQL, xCon
    If Rstabo.State = 0 Then GoTo Salir
    '-----------------------------------------------------------
        
        '--------------
        rst.MoveFirst
        For A = 1 To rst.RecordCount    '--GRUPO DE CLIENTE/PROVEEDOR
            DoEvents
            '--SI SE NTERRUMPE EL PROCESO => SALIR
            If BAND_INTERRUMPIR = True Then GoTo Salir:
            ProgressBar1.Value = A
            
            xSaldoDoc = 0
            
            If NulosC(rst("nombre")) <> xNomPro Then
                DoEvents
                Cambio = True
                xNomPro = NulosC(rst("nombre"))
                Fg1.Rows = Fg1.Rows + 1
                xFila = xFila + 1
                Fg1.TextMatrix(xFila, 4) = "TOTAL -->"
                
                '--acumulando los totales por grupo
                TotDebe = NulosN(GRID_SUMAR_COL(Fg1, 10, xFilaIniGrupo, Fg1.Rows - 2))
                TotHaber = NulosN(GRID_SUMAR_COL(Fg1, 11, xFilaIniGrupo, Fg1.Rows - 2))
                
                Fg1.TextMatrix(xFila, 10) = Format(TotDebe, FORMAT_MONTO)
                Fg1.TextMatrix(xFila, 11) = Format(TotHaber, FORMAT_MONTO)
                
''''                If OptCliente.Value = True Then
''''                    Fg1.TextMatrix(xFila, 12) = Format(TotDebe - TotHaber, FORMAT_MONTO)
''''                Else
''''                    Fg1.TextMatrix(xFila, 12) = Format(TotHaber - TotDebe, FORMAT_MONTO)
''''                End If
                
                Select Case xIdLibro
                    Case 1, 4, 9, 40, 999
                        Fg1.TextMatrix(xFila, 12) = Format(TotHaber - TotDebe, FORMAT_MONTO)
                    Case 2, 37, 40, 41
                        Fg1.TextMatrix(xFila, 12) = Format(TotDebe - TotHaber, FORMAT_MONTO)
                End Select
                
                '*****resumen
                Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(TotDebe, FORMAT_MONTO)
                Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(TotHaber, FORMAT_MONTO)
                
''''                If OptCliente.Value = True Then
''''                    Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotDebe - TotHaber, FORMAT_MONTO)
''''                Else
''''                    Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotHaber - TotDebe, FORMAT_MONTO)
''''                End If
                Select Case xIdLibro
                    Case 1, 4, 9, 40, 999
                        Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotHaber - TotDebe, FORMAT_MONTO)
                    Case 2, 37, 40, 41
                        Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotDebe - TotHaber, FORMAT_MONTO)
                End Select
                
                
                '******
                
                With Fg1
                    .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H80000008
                    .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
                    .FillStyle = flexFillRepeat
                    .CellFontBold = True
                End With
                
                '----MOSTRAR SOLO DESCUADRADOS ---------
                If chk_descuadrado.Value = 1 Then
                    If NulosN(Fg1.TextMatrix(xFila, 12)) = 0 Then
                        GRID_DELETE Fg1, Fg1.Rows - 2, Fg1.Rows - 1, e_Fila
                        Fg1.Rows = Fg1.Rows + 1
                        xFila = Fg1.Rows - 1
                    Else
                        Fg1.Rows = Fg1.Rows + 2
                        xFila = xFila + 2
                    End If
                    '---del resumen
                    If NulosN(Fg2.TextMatrix(Fg2.Rows - 1, 5)) = 0 Then
                        GRID_DELETE Fg2, Fg2.Rows - 1, Fg2.Rows - 1, e_Fila
                    End If
                    '---------------
                Else
                    Fg1.Rows = Fg1.Rows + 2
                    xFila = xFila + 2
                End If
                '---------------------------------------------------------
                TotGralHaber = TotGralHaber + TotHaber
                TotGralDebe = TotGralDebe + TotDebe
                
                TotHaber = 0
                TotDebe = 0
                '---------------------------------------------------------

                GRID_COMBINAR Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 12, "Nº R.U.C. : " & RellenarBlancos(NulosC(rst("numruc")), 12, 1) & "  " & xNomPro, flexAlignLeftCenter, True, , , , True
                xFilaIniGrupo = Fg1.Rows - 1
                
                '*****resumen
                Fg2.Rows = Fg2.Rows + 1
                Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(rst("numruc"))
                Fg2.TextMatrix(Fg2.Rows - 1, 2) = xNomPro
                '******

                
                With Fg1
                    .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H800000
                    .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
                    .FillStyle = flexFillRepeat
                    .CellFontBold = True
                End With
            Else
                Cambio = False
            End If

            Fg1.Rows = Fg1.Rows + 1
            xFila = xFila + 1
            xFilaIni = xFila
            
            Fg1.TextMatrix(xFila, 1) = NulosC(rst("registro"))
            
            Fg1.TextMatrix(xFila, 2) = NulosC(rst("libro"))
            Fg1.TextMatrix(xFila, 3) = NulosC(rst("abrev"))
            Fg1.TextMatrix(xFila, 4) = NulosC(rst("numdoc2"))
            Fg1.TextMatrix(xFila, 5) = Format(rst("fchdoc"), FORMAT_DATE)
            Fg1.TextMatrix(xFila, 6) = Format(rst("fchven"), FORMAT_DATE)
            Fg1.TextMatrix(xFila, 7) = NulosC(rst("simbolo"))
            Fg1.TextMatrix(xFila, 8) = Format(NulosN(rst("imptotal")), FORMAT_MONTO)
            Fg1.TextMatrix(xFila, 9) = Format(NulosN(rst("tipcam")), "###0.##0") & ""
            xOrigen = NulosC(rst("xOrigen"))
            If NulosC(rst("xOrigen")) = "D" Then
                Fg1.TextMatrix(xFila, 10) = Format(NulosN(rst(nCampoMuestra)), FORMAT_MONTO)
                TotDebe = TotDebe + NulosN(rst(nCampoMuestra))
            Else
                Fg1.TextMatrix(xFila, 11) = Format(NulosN(rst(nCampoMuestra)), FORMAT_MONTO) '--saldo
                TotHaber = TotHaber + NulosN(rst(nCampoMuestra))
            End If
            Fg1.TextMatrix(xFila, 12) = Format(NulosN(rst(nCampoMuestra)), FORMAT_MONTO)
            
            Fg1.TextMatrix(xFila, 13) = NulosC(rst("numdoc2"))
            Fg1.TextMatrix(xFila, 14) = NulosC(rst("glosaope"))
            
            xSaldoDoc = NulosN(rst("impsal"))
            
            '--filtrar los movimientos de las provisiones para proceder a obtener el saldo actual
            Rstabo.Filter = "rregistro = '" & NulosC(rst("registro")) & "' and iddoc= " & NulosN(rst("id"))
            '-------------------------------------------------------------
            If Rstabo.RecordCount <> 0 Then
                Rstabo.MoveFirst
                '--ordenar el rst para mostrar el detalle
                '--NOta: el primer orden es importante pues indica que el ajuste por diferencia de cambio se mostrara en la ultima posicion del detalle,
                '--      esto indicara que se esta aplicando un ajuste al documento para mostrar el saldo a cero.
                Rstabo.Sort = "libro desc,fchemi ASC"
                
                Do While Not Rstabo.EOF
                    '--SI SE NTERRUMPE EL PROCESO => SALIR
'                    DoEvents
                    If BAND_INTERRUMPIR = True Then GoTo Salir:
                    Fg1.Rows = Fg1.Rows + 1
                    xFila = xFila + 1
                    
                    Fg1.TextMatrix(xFila, 1) = NulosC(Rstabo("registro"))
                    Fg1.TextMatrix(xFila, 2) = NulosC(Rstabo("libro"))
                    Fg1.TextMatrix(xFila, 3) = NulosC(Rstabo("abrev"))
                    Fg1.TextMatrix(xFila, 4) = NulosC(Rstabo("numdoc"))
                    Fg1.TextMatrix(xFila, 5) = Format(Rstabo("fchemi"), FORMAT_DATE)
                    Fg1.TextMatrix(xFila, 7) = NulosC(Rstabo("simbolo"))
                    Fg1.TextMatrix(xFila, 8) = Format(NulosN(Rstabo("imptotal")), FORMAT_MONTO)
                    Fg1.TextMatrix(xFila, 9) = Format(NulosN(Rstabo("tipcam")), "####.###")
                    
''                    '--verificar si el libro es de ajuste por dif de cambio
''                    If InStr(LCase(Rstabo("libro")), "ajuste") <> 0 Then
''
''                        '--verificar si importe estaba en el debe
''                        If NulosN(Rstabo("impdeb")) = 0 Then
''                            Fg1.TextMatrix(xFila, 10) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
''                            TotDebe = TotDebe + NulosN(Rstabo(nCampoMuestra))
''                            If xOrigen = "D" Then
''                                Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) + NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
''                            Else
''                                Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) - NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
''                            End If
''
''                        Else
''                            Fg1.TextMatrix(xFila, 11) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
''                            TotHaber = TotHaber + NulosN(Rstabo(nCampoMuestra))
''                            If xOrigen = "D" Then
''                                Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) - NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
''                            Else
''                                Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) + NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
''                            End If
''                        End If
''
''                    Else
                    
                        '--verificar si importe estaba en el debe
                        If NulosN(Rstabo("impdeb")) <> 0 Then
                            Fg1.TextMatrix(xFila, 10) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                            TotDebe = TotDebe + NulosN(Rstabo(nCampoMuestra))
                            If xOrigen = "D" Then
                                Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) + NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                            Else
                                Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) - NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                            End If
                            
                        Else
                            Fg1.TextMatrix(xFila, 11) = Format(NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                            TotHaber = TotHaber + NulosN(Rstabo(nCampoMuestra))
                            If xOrigen = "D" Then
                                Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) - NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                            Else
                                Fg1.TextMatrix(xFila, 12) = Format(NulosN(Fg1.TextMatrix(xFila - 1, 12)) + NulosN(Rstabo(nCampoMuestra)), FORMAT_MONTO)
                            End If
                        End If
                    
''                    End If
                    
                    Fg1.TextMatrix(xFila, 13) = NulosC(Rstabo("numdoc2"))
                    Fg1.TextMatrix(xFila, 14) = NulosC(Rstabo("rglosaope"))
                                        
                    Rstabo.MoveNext
                Loop
                
            End If
            
            '---ACTUALIZANDO EL SALDO AL DOCUMENTO
            '--solo se actualizara el saldo si el documento esta en la moneda de consulta
            '--considerar actualizar el saldo si es ajuste por diferencia de cambio sin importar la moneda
            If (xSaldoDoc <> NulosN(Fg1.TextMatrix(xFila, 12)) And NulosN(rst("idmon")) = NulosN(TxtIdMon.Text)) Or InStr(LCase(Fg1.TextMatrix(xFila, 2)), "ajuste") <> 0 Then
                
                '--obtener el ultimo saldo del documento
                If InStr(LCase(Fg1.TextMatrix(xFila, 2)), "ajuste") <> 0 Then
                    sSaldoFinal = 0
                Else
                    sSaldoFinal = NulosN(Fg1.TextMatrix(xFila, 12))
                End If
                '--------------------------------------------------------
                
'''                If OptCliente.Value = True Then     '--VENTAS
'''                    xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = " & sSaldoFinal & " WHERE (((vta_ventas.id)=" & rst("id") & "))"
'''
'''                ElseIf OptProvee.Value = True Then                                '--COMPRAS
'''                    If LCase(NulosC(rst("libro"))) = "percepciones" Then
'''                        xCon.Execute "UPDATE con_percepcion SET con_percepcion.impsal = " & sSaldoFinal & " WHERE (((con_percepcion.id)=" & rst("id") & "))"
'''                    Else
'''                        xCon.Execute "UPDATE com_compras SET com_compras.impsal = " & sSaldoFinal & " WHERE (((com_compras.id)=" & rst("id") & "))"
'''                    End If
'''
'''                ElseIf opt4ta.Value = True Then
'''                    xCon.Execute "UPDATE com_honorarios SET com_honorarios.impsal = " & sSaldoFinal & " WHERE (((com_honorarios.id)=" & rst("id") & "))"
'''
'''                End If
                
        
                If sSaldoFinal <> NulosN(rst("impsal")) Then
                
                    '--Actualizar saldos a documento
                    Select Case xIdLibro
                        Case 1   '--Compras
                            xCon.Execute "Update com_compras set com_compras.impsal=" & sSaldoFinal & " where com_compras.id = " & NulosC(rst("id"))
                        Case 4 '--Percepcion
                            xCon.Execute "Update con_percepcion set con_percepcion.impsal=" & sSaldoFinal & " where con_percepcion.id = " & NulosC(rst("id"))
                        Case 40  '--Honorarios
                            xCon.Execute "Update com_honorarios set com_honorarios.impsal=" & sSaldoFinal & " where com_honorarios.id = " & NulosC(rst("id"))
                        Case 2   '--Ventas
                            xCon.Execute "Update vta_ventas set vta_ventas.impsal=" & sSaldoFinal & " where vta_ventas.id = " & NulosC(rst("id"))
                        Case 9  '--Boleta Pago
                            xCon.Execute "Update pla_boleta set pla_boleta.impsal=" & sSaldoFinal & " where pla_boleta.id = " & NulosC(rst("id"))
                        Case 37  '--Letras
                            xCon.Execute "Update let_letradet set let_letradet.impsal=" & sSaldoFinal & " where let_letradet.corr = " & NulosC(rst("id"))
                        Case 41  '--Lgd, Lgc
                            xCon.Execute "Update vta_gastodebito set vta_gastodebito.impsal=" & sSaldoFinal & " where vta_gastodebito.id = " & NulosC(rst("id"))
                        Case 42  '--Planilla letras
                            xCon.Execute "Update let_planilla set let_planilla.impsal=" & sSaldoFinal & " where let_planilla.id = " & NulosC(rst("id"))
                        Case 999 '--Reembolsables
                            xCon.Execute "Update com_reembolsables set com_reembolsables.impsal=" & sSaldoFinal & " where com_reembolsables.id = " & NulosC(rst("id"))
                    End Select
                
                End If
                
                
            End If
            
            '----MOSTRAR SOLO DESCUADRADOS ---------
            If chk_descuadrado.Value = 1 And OptTodos.Value = True Then
                If NulosN(Fg1.TextMatrix(xFila, 12)) >= 0 Then
                'TabOne1.CurrTab = 0
                    GRID_DELETE Fg1, Fg1.Rows - 1 - Rstabo.RecordCount, Fg1.Rows - 1, e_Fila
                    '*********************************************
                    If Rstabo.RecordCount <> 0 Then
                        Rstabo.MoveFirst
                        Do While Not Rstabo.EOF
'''                            If OptCliente.Value = True Then
'''                                TotHaber = TotHaber - NulosN(Rstabo(nCampoMuestra))
'''                            Else
'''                                TotDebe = TotDebe - NulosN(Rstabo(nCampoMuestra))
'''                            End If
                            Select Case xIdLibro
                                Case 1, 4, 9, 40, 999
                                    TotDebe = TotDebe - NulosN(Rstabo(nCampoMuestra))
                                Case 2, 37, 40, 41
                                    TotHaber = TotHaber - NulosN(Rstabo(nCampoMuestra))
                            End Select
                            Rstabo.MoveNext
                        Loop
                    End If
                    
'*********************************************
' Modificado 29/03/12 - Jose Chacon - Cometar Lineas
'*********************************************
'*********************************************************************************
'                    Select Case xIdLibro
'                        Case 1, 4, 9, 40, 999
'                            TotDebe = TotDebe - NulosN(Rstabo(nCampoMuestra))
'                        Case 2, 37, 40, 41
'                            TotHaber = TotHaber - NulosN(Rstabo(nCampoMuestra))
'                    End Select
'*********************************************************************************

                            
'''                    If OptCliente.Value = True Then
'''                        TotDebe = TotDebe - NulosN(rst(nCampoMuestra))
'''                    Else
'''                        TotHaber = TotHaber - NulosN(rst(nCampoMuestra))
'''                    End If
                    
                    '*********************************************
                    xFila = Fg1.Rows - 1
                    mRowIni = -1
                Else
                    mRowIni = 0
                End If
            End If
            '---------------------------------------------------------
            
            rst.MoveNext
            If rst.EOF = True Then
                Exit For
            End If
            If mRowIni = 0 Then
                If xColor = 0 Then
                    GRID_COLOR_FONDO Fg1, xFilaIni, 1, Fg1.Rows - 1, Fg1.Cols - 1, &H80000005
                    xColor = 1
                Else
                    GRID_COLOR_FONDO Fg1, xFilaIni, 1, Fg1.Rows - 1, Fg1.Cols - 1, &HE0DCDA
                    xColor = 0
                End If
            End If
            

        Next A
        
        
        
        
        Fg1.Rows = Fg1.Rows + 1
        xFila = xFila + 1
        Fg1.TextMatrix(xFila, 4) = "TOTAL -->"


        '--acumulando los totales por grupo
        TotDebe = NulosN(GRID_SUMAR_COL(Fg1, 10, xFilaIniGrupo, Fg1.Rows - 2))
        TotHaber = NulosN(GRID_SUMAR_COL(Fg1, 11, xFilaIniGrupo, Fg1.Rows - 2))

        '---------------------------------------------------------
        TotGralHaber = TotGralHaber + TotHaber
        TotGralDebe = TotGralDebe + TotDebe
        '---------------------------------------------------------
        
        Fg1.TextMatrix(xFila, 10) = Format(TotDebe, FORMAT_MONTO)
        Fg1.TextMatrix(xFila, 11) = Format(TotHaber, FORMAT_MONTO)
                
''''        If OptCliente.Value = True Then
''''            Fg1.TextMatrix(xFila, 12) = Format(TotDebe - TotHaber, FORMAT_MONTO)
''''        Else
''''            Fg1.TextMatrix(xFila, 12) = Format(TotHaber - TotDebe, FORMAT_MONTO)
''''        End If
        
        Select Case xIdLibro
            Case 1, 4, 9, 40, 999
                Fg1.TextMatrix(xFila, 12) = Format(TotHaber - TotDebe, FORMAT_MONTO)
            Case 2, 37, 40, 41
                Fg1.TextMatrix(xFila, 12) = Format(TotDebe - TotHaber, FORMAT_MONTO)
        End Select
        '*****resumen
        Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(TotDebe, FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(TotHaber, FORMAT_MONTO)
        
''''        If OptCliente.Value = True Then
''''            Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotDebe - TotHaber, FORMAT_MONTO)
''''        Else
''''            Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotHaber - TotDebe, FORMAT_MONTO)
''''        End If
        
        Select Case xIdLibro
            Case 1, 4, 9, 40, 999
                Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotHaber - TotDebe, FORMAT_MONTO)
            Case 2, 37, 40, 41
                Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotDebe - TotHaber, FORMAT_MONTO)
        End Select
        
        '******

        With Fg1
            .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H80000008
            .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
            .FillStyle = flexFillRepeat
            .CellFontBold = True
        End With
        '----MOSTRAR SOLO DESCUADRADOS ---------
        If chk_descuadrado.Value = 1 Then
            If NulosN(Fg1.TextMatrix(xFila, 12)) = 0 Then
                GRID_DELETE Fg1, Fg1.Rows - 2, Fg1.Rows - 1, e_Fila
                xFila = Fg1.Rows - 1
            End If
            '--del resumen
            If NulosN(Fg2.TextMatrix(Fg2.Rows - 1, 5)) = 0 Then
                GRID_DELETE Fg2, Fg2.Rows - 1, Fg2.Rows - 1, e_Fila
            End If
            '---------------
        End If
        
        '---------------------------------------------------------

        If TotGralDebe <> 0 Or TotGralHaber <> 0 Then
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = "TOTAL GRAL -->"
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(TotGralDebe, FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(TotGralHaber, FORMAT_MONTO)
            
'''''            If OptCliente.Value = True Then
'''''                Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(TotGralDebe - TotGralHaber, FORMAT_MONTO)
'''''            Else
'''''                Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(TotGralHaber - TotGralDebe, FORMAT_MONTO)
'''''            End If
            
            Select Case xIdLibro
                Case 1, 4, 9, 40, 999
                    Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(TotGralHaber - TotGralDebe, FORMAT_MONTO)
                Case 2, 37, 40, 41
                    Fg1.TextMatrix(Fg1.Rows - 1, 12) = Format(TotGralDebe - TotGralHaber, FORMAT_MONTO)
            End Select
            
            With Fg1
                .Cell(flexcpForeColor, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1) = &H80000008
                .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1
                .FillStyle = flexFillRepeat
                .CellFontBold = True
            End With
            '*****resumen
            Fg2.Rows = Fg2.Rows + 2
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = "TOTAL GRAL -->"
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(TotGralDebe, FORMAT_MONTO)
            Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(TotGralHaber, FORMAT_MONTO)
            
'''''            If OptCliente.Value = True Then
'''''                Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotGralDebe - TotGralHaber, FORMAT_MONTO)
'''''            Else
'''''                Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotGralHaber - TotGralDebe, FORMAT_MONTO)
'''''            End If
            Select Case xIdLibro
                Case 1, 4, 9, 40, 999
                    Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotGralHaber - TotGralDebe, FORMAT_MONTO)
                Case 2, 37, 40, 41
                    Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(TotGralDebe - TotGralHaber, FORMAT_MONTO)
            End Select
            With Fg2
                .Cell(flexcpForeColor, Fg2.Rows - 1, 1, Fg2.Rows - 1, Fg2.Cols - 1) = &H80000008
                .Select Fg2.Rows - 1, 1, Fg2.Rows - 1, Fg2.Cols - 1
                .FillStyle = flexFillRepeat
                .CellFontBold = True
            End With
            '******
    End If
    
    End If
    If mRowIni = 0 Then
        If xColor = 0 Then
            GRID_COLOR_FONDO Fg1, xFilaIni, 1, Fg1.Rows - 2, Fg1.Cols - 1, &H80000005
            xColor = 1
        Else
            GRID_COLOR_FONDO Fg1, xFilaIni, 1, Fg1.Rows - 2, Fg1.Cols - 1, &HE0DCDA
            xColor = 0
        End If
    End If
    
    '--ajustar totales
    '--detalle
    Fg1.AutoSizeMode = flexAutoSizeColWidth
    Fg1.AutoSize 8
    Fg1.AutoSize 10
    Fg1.AutoSize 11
    Fg1.AutoSize 12
    '--resumen
    Fg2.AutoSizeMode = flexAutoSizeColWidth
    Fg2.AutoSize 3
    Fg2.AutoSize 4
    Fg2.AutoSize 5
    '----------------------------------------------------------------------
    
    Set rst = Nothing
    Set Rstabo = Nothing
    fraBarra.Visible = False
    Me.MousePointer = vbDefault
''    MsgBox "La Consulta fue se realizó Correctamente", vbInformation, xTitulo
    Exit Sub
    
Salir:
'''''    Resume
    Me.MousePointer = vbDefault
    fraBarra.Visible = False
    Set rst = Nothing
    Set Rstabo = Nothing
    MsgBox "La Consulta fue Interrumpida", vbInformation, xTitulo
    Exit Sub
error:
    
    Me.MousePointer = vbDefault
    fraBarra.Visible = False
    Set rst = Nothing
    Set Rstabo = Nothing
    SHOW_ERROR Me.Name, "CargarCli2"
    
End Sub

Sub PreparaRST()
    '===================================================================================================
    'Creado: 11/01/12 Johan Castro
    'Propósito: Definir Rst Temporal para totalizar
    '
    'Entradas:  Ninguno
    '
    'Resultados: Rst Temporal definido, listo para ser usado
    '===================================================================================================
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(8, 3) As String

    xCampos(0, 0) = "ruc":          xCampos(0, 1) = "C":      xCampos(0, 2) = "20" ' ruc
    xCampos(1, 0) = "nombre":       xCampos(1, 1) = "C":      xCampos(1, 2) = "240" ' nombre
    xCampos(2, 0) = "impmn":        xCampos(2, 1) = "D":      xCampos(2, 2) = "2" ' importe total del documento
    xCampos(3, 0) = "impme":        xCampos(3, 1) = "D":      xCampos(3, 2) = "2" ' saldo del documento
    xCampos(4, 0) = "pimpmn":       xCampos(4, 1) = "D":      xCampos(4, 2) = "2" ' saldo del documento en la moneda de trabajo
    xCampos(5, 0) = "pimpme":       xCampos(5, 1) = "D":      xCampos(5, 2) = "2" ' importe acuenta
    xCampos(6, 0) = "simpmn":       xCampos(6, 1) = "D":      xCampos(6, 2) = "2" ' nuevo saldo del documento
    xCampos(7, 0) = "simpme":       xCampos(7, 1) = "D":      xCampos(7, 2) = "2" ' nuevo saldo del documento

    Set xRstTot = xFun.CrearRstTMP(xCampos)
    xRstTot.Open
End Sub

Private Sub Acumular(xruc As String, xnombre As String, xmn As Double, xme As Double, xpmn As Double, xpme As Double)
    '===================================================================================================
    'Creado: 11/01/12 Johan Castro
    'Propósito: Acumular importes por proveedor/cliente
    '
    'Entradas:  xruc=Ruc del cliente, proveedor
    '           xnombre=Razon social
    '           xmn=Provicion en MN
    '           xme=Provicion en ME
    '           xpmn=Pagos/Cobranza en MN
    '           xpme=Pagos/Cobranza en ME
    'Resultados: Rst con datos agrupados segun ruc
    '
    '===================================================================================================
    xRstTot.Filter = ""
    xRstTot.Filter = "ruc='" & xruc & "'"
    If xRstTot.RecordCount = 0 Then
        xRstTot.AddNew
    End If

    xRstTot("ruc") = xruc
    xRstTot("nombre") = xnombre
    xRstTot("impmn") = NulosN(xRstTot("impmn")) + xmn
    xRstTot("impme") = NulosN(xRstTot("impme")) + xme
    xRstTot("pimpmn") = NulosN(xRstTot("pimpmn")) + xpmn
    xRstTot("pimpme") = NulosN(xRstTot("pimpme")) + xpme
End Sub



Private Function BuscarNCBancos() As String
    Dim nSQL As String
    Dim nSQLWhere As String
    Dim nSQLPer As String
    Dim nSQLFecha As String
    Dim nSQLApertura As String
    Dim xRstNC As New ADODB.Recordset

    '--Si libro es distinto a compras o ventas, salir de evento
    If xIdLibro <> 1 And xIdLibro <> 2 Then Exit Function

    If OptFch(0).Value = True Then '--x fecha de documento
        nSQLFecha = " and ( vta_ventas.fchdoc between CDate('" & TxtFchIni.Valor & "') and CDate('" & TxtFchFin.Valor & "') )"
    ElseIf OptFch(1).Value = True Then '--x fecha de registro
        nSQLFecha = " and ( vta_ventas.fchreg between CDate('" & TxtFchIni.Valor & "') and CDate('" & TxtFchFin.Valor & "') )"
    End If
    
'    If NulosN(LblIdCliPro.Caption) <> 0 Then
        nSQLPer = " and vta_ventas.idcli =" & NulosN(LblIdCliPro.Caption)
'        If OptAperturaCon.Value = True Then nSQLApertura = " or (vta_ventas.numreg='000001' " & nSQLPer & " ) "
'    Else
'        If OptAperturaCon.Value = True Then nSQLApertura = " or vta_ventas.numreg='000001' "
'    End If
'    If OptAperturaSin.Value = True Then nSQLApertura = " and vta_ventas.numreg<>'000001' "
'    If OptAperturaSolo.Value = True Then nSQLApertura = " and vta_ventas.numreg='000001' "
    '--------------------------------------------
    nSQLWhere = nSQLPer
    '--------------------------------------------

    nSQL = " SELECT tes_cajaorigendet.iddoc " _
        + vbCr + " FROM tes_cajaorigendet INNER JOIN vta_ventas ON tes_cajaorigendet.iddoc = vta_ventas.id " _
        + vbCr + " WHERE ((tes_cajaorigendet.idmod=2) AND ((vta_ventas.tipdoc)=7)) " & nSQLWhere _
        + vbCr + " UNION " _
        + vbCr + " SELECT tes_cajadestinodet.iddoc " _
        + vbCr + " FROM tes_cajadestinodet INNER JOIN vta_ventas ON tes_cajadestinodet.iddoc = vta_ventas.id " _
        + vbCr + " WHERE ((tes_cajadestinodet.idmod=2) AND ((vta_ventas.tipdoc)=7)) " & nSQLWhere
    
    If xIdLibro = 1 Then
        nSQL = Replace(nSQL, ".idmod=2", ".idmod=1")
        nSQL = Replace(nSQL, "vta_ventas.idcli", "com_compras.idpro")
        nSQL = Replace(nSQL, "vta_ventas", "com_compras")
    End If

    RST_Busq xRstNC, nSQL, xCon
    If xRstNC.State = 1 Then
        If xRstNC.RecordCount <> 0 Then
            BuscarNCBancos = RstRegistroGenerarId(xRstNC, "iddoc", "", "", True)
        End If
    End If
    Set xRstNC = Nothing


End Function
