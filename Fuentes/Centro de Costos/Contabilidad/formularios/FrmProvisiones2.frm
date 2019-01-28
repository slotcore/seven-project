VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmProvisiones2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Asientos Diversos"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   720
      Left            =   12120
      TabIndex        =   74
      Top             =   3900
      Visible         =   0   'False
      Width           =   5805
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   270
         Left            =   150
         TabIndex        =   75
         Top             =   375
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   2
         X1              =   0
         X2              =   5820
         Y1              =   15
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   3
         X1              =   15
         X2              =   15
         Y1              =   30
         Y2              =   1170
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   2
         X1              =   5790
         X2              =   5790
         Y1              =   15
         Y2              =   1155
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   5835
         Y1              =   705
         Y2              =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Importando..."
         Height          =   195
         Left            =   195
         TabIndex        =   76
         Top             =   150
         Width           =   930
      End
   End
   Begin VB.Frame Frame7 
      BorderStyle     =   0  'None
      Caption         =   "2"
      Height          =   5400
      Left            =   11640
      TabIndex        =   51
      Top             =   6300
      Visible         =   0   'False
      Width           =   11610
      Begin VB.CommandButton CmdBusMod 
         Height          =   240
         Left            =   1785
         Picture         =   "FrmProvisiones2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   435
         Width           =   255
      End
      Begin VB.TextBox TxtIdMod 
         Height          =   300
         Left            =   1275
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   67
         Text            =   "TxtIdMod"
         Top             =   405
         Width           =   795
      End
      Begin VB.OptionButton OptHaber 
         Caption         =   "Haber"
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
         Left            =   2265
         TabIndex        =   63
         Top             =   750
         Width           =   900
      End
      Begin VB.OptionButton OptDebe 
         Caption         =   "Debe"
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
         Left            =   1275
         TabIndex        =   62
         Top             =   750
         Width           =   900
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg2 
         Height          =   3645
         Left            =   45
         TabIndex        =   52
         Top             =   1035
         Width           =   11490
         _cx             =   20267
         _cy             =   6429
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
         Rows            =   50
         Cols            =   23
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmProvisiones2.frx":0132
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
      Begin VB.Frame Frame8 
         Height          =   720
         Left            =   45
         TabIndex        =   53
         Top             =   4620
         Width           =   7200
         Begin VB.CommandButton CmdDelTodo 
            Caption         =   "&Eliminar Todo"
            Height          =   345
            Left            =   2790
            TabIndex        =   61
            Top             =   240
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.CommandButton Command3 
            Caption         =   "E&xportar"
            Height          =   345
            Left            =   5610
            TabIndex        =   60
            Top             =   240
            Width           =   1200
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Importar"
            Height          =   345
            Left            =   4380
            TabIndex        =   59
            Top             =   240
            Width           =   1200
         End
         Begin VB.CommandButton CmdDel2 
            Caption         =   "&Eliminar"
            Height          =   345
            Left            =   1560
            TabIndex        =   56
            Top             =   240
            Width           =   1200
         End
         Begin VB.CommandButton CmdAdd2 
            Caption         =   "&Agregar"
            Height          =   345
            Left            =   330
            TabIndex        =   55
            Top             =   240
            Width           =   1200
         End
      End
      Begin VB.Frame Frame9 
         Height          =   720
         Left            =   7335
         TabIndex        =   57
         Top             =   4620
         Width           =   4200
         Begin VB.CommandButton CmdAcepta 
            Caption         =   "&Aceptar"
            Height          =   345
            Left            =   1455
            TabIndex        =   58
            Top             =   240
            Width           =   1200
         End
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   11595
         X2              =   11565
         Y1              =   15
         Y2              =   5385
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   11580
         Y1              =   5385
         Y2              =   5385
      End
      Begin VB.Label LblDescModulo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblDescModulo"
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
         Left            =   2100
         TabIndex        =   68
         Top             =   405
         Width           =   4020
      End
      Begin VB.Label Label6 
         Caption         =   "Modulo"
         Height          =   195
         Left            =   135
         TabIndex        =   65
         Top             =   435
         Width           =   1050
      End
      Begin VB.Label Label3 
         Caption         =   "Naturaleza"
         Height          =   195
         Left            =   135
         TabIndex        =   64
         Top             =   735
         Width           =   1050
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   11580
         Y1              =   15
         Y2              =   0
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   5370
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de la Cuenta"
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
         Left            =   225
         TabIndex        =   54
         Top             =   75
         Width           =   1755
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   45
         Top             =   45
         Width           =   11520
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   0
      TabIndex        =   12
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
         Caption         =   "Detalle de la Cuenta"
         Height          =   6795
         Left            =   12525
         TabIndex        =   16
         Top             =   375
         Width           =   11790
         Begin VB.CheckBox ChkAjusteDifCambio 
            Caption         =   "Aplica Ajuste x Dif. Cambio"
            Enabled         =   0   'False
            Height          =   315
            Left            =   9450
            TabIndex        =   77
            Top             =   1980
            Width           =   2415
         End
         Begin VB.TextBox TxtTC 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   7575
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   6
            Text            =   "TxtTC"
            Top             =   1950
            Width           =   900
         End
         Begin VB.CommandButton CmdBusLib 
            Enabled         =   0   'False
            Height          =   240
            Left            =   2220
            Picture         =   "FrmProvisiones2.frx":03DE
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   420
            Width           =   210
         End
         Begin VB.CommandButton cb 
            Height          =   240
            Index           =   1
            Left            =   2220
            Picture         =   "FrmProvisiones2.frx":0510
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   765
            Width           =   210
         End
         Begin VB.Frame Frame6 
            Caption         =   "[ Tipo de Cambio ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   690
            Left            =   9450
            TabIndex        =   39
            Top             =   960
            Width           =   2250
            Begin VB.Label LblTipoCambio 
               Alignment       =   2  'Center
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
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   120
               TabIndex        =   40
               Top             =   300
               Width           =   1980
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "[ Periodo ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   720
            Left            =   9450
            TabIndex        =   33
            Top             =   195
            Width           =   2250
            Begin VB.Label lblperiodo 
               Alignment       =   2  'Center
               Caption         =   "lblperiodo"
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
               TabIndex        =   34
               Top             =   330
               Width           =   1995
            End
         End
         Begin VB.CommandButton cb 
            Height          =   240
            Index           =   0
            Left            =   2220
            Picture         =   "FrmProvisiones2.frx":0642
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   1080
            Width           =   210
         End
         Begin VB.CommandButton CmdBusPro 
            Height          =   240
            Left            =   8400
            Picture         =   "FrmProvisiones2.frx":0774
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   1260
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2640
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   5
            Text            =   "TxtNumDoc"
            Top             =   1995
            Width           =   1830
         End
         Begin VB.TextBox TxtSerDoc 
            Height          =   300
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   4
            Text            =   "TxtSerDoc"
            Top             =   1995
            Width           =   900
         End
         Begin VB.CommandButton CmdBusTipDoc 
            Height          =   240
            Left            =   2220
            Picture         =   "FrmProvisiones2.frx":08A6
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   1710
            Width           =   210
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   3630
            Left            =   195
            TabIndex        =   8
            Top             =   2640
            Width           =   9870
            _cx             =   17410
            _cy             =   6403
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
            Rows            =   50
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmProvisiones2.frx":09D8
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
         Begin VB.Frame Frame3 
            Height          =   3720
            Left            =   10170
            TabIndex        =   19
            Top             =   2550
            Width           =   1560
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   540
               Top             =   300
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.CommandButton CmdImportar 
               Caption         =   "Importar"
               Height          =   690
               Left            =   120
               TabIndex        =   73
               Top             =   2910
               Width           =   1305
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Agregar Documentos"
               Height          =   690
               Left            =   1320
               TabIndex        =   41
               Top             =   300
               Visible         =   0   'False
               Width           =   1305
            End
            Begin VB.CommandButton CmdAdd 
               Caption         =   "Agregar Cuenta"
               Height          =   690
               Left            =   120
               TabIndex        =   21
               Top             =   960
               Width           =   1305
            End
            Begin VB.CommandButton CmdDel 
               Caption         =   "Eliminar Cuenta"
               Height          =   690
               Left            =   120
               TabIndex        =   20
               Top             =   1665
               Width           =   1305
            End
         End
         Begin VB.TextBox TxtGlosa 
            Height          =   300
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "TxtGlosa"
            Top             =   2310
            Width           =   10125
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmi 
            Height          =   300
            Left            =   1560
            TabIndex        =   2
            Top             =   1365
            Width           =   1290
            _ExtentX        =   2275
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
            Valor           =   "17/07/2008"
         End
         Begin VB.TextBox TxtTipDoc 
            Height          =   300
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   3
            Text            =   "TxtTipDoc"
            Top             =   1680
            Width           =   900
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   0
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   1
            Text            =   "txt_cb(0)"
            Top             =   1050
            Width           =   900
         End
         Begin VB.TextBox TxtNombre 
            Height          =   300
            Left            =   7575
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   11
            Text            =   "TxtNombre"
            Top             =   1230
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Frame Frame5 
            Height          =   570
            Left            =   195
            TabIndex        =   17
            Top             =   6225
            Width           =   11505
            Begin VB.TextBox TxtTotDebDol 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   300
               Left            =   7650
               Locked          =   -1  'True
               TabIndex        =   70
               Text            =   "TxtTotDebDol"
               Top             =   180
               Width           =   960
            End
            Begin VB.TextBox TxtTotHabDol 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   300
               Left            =   8610
               Locked          =   -1  'True
               TabIndex        =   69
               Text            =   "TxtTotHabDol"
               Top             =   180
               Width           =   960
            End
            Begin VB.TextBox TxtTotHab 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   300
               Left            =   6690
               Locked          =   -1  'True
               TabIndex        =   10
               Text            =   "TxtTotHab"
               Top             =   180
               Width           =   960
            End
            Begin VB.TextBox TxtTotDeb 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   300
               Left            =   5730
               Locked          =   -1  'True
               TabIndex        =   9
               Text            =   "TxtTotDeb"
               Top             =   180
               Width           =   960
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total ==>"
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
               Left            =   4800
               TabIndex        =   18
               Top             =   210
               Width           =   825
            End
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   1
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   0
            Text            =   "txt_cb(1)"
            Top             =   735
            Width           =   900
         End
         Begin VB.TextBox TxtIdLibro 
            BackColor       =   &H8000000F&
            Height          =   300
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   47
            Text            =   "TxtIdLibro"
            Top             =   390
            Width           =   900
         End
         Begin VB.Label lblReg 
            Caption         =   "lblReg"
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
            Height          =   270
            Left            =   9360
            TabIndex        =   72
            Top             =   1680
            Width           =   2250
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "T.C."
            Height          =   195
            Index           =   0
            Left            =   6990
            TabIndex        =   71
            Top             =   2040
            Width           =   300
         End
         Begin VB.Label lbl_cb_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb_cod(0)"
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
            Left            =   4680
            TabIndex        =   30
            Top             =   1050
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.Label LblLibro 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblLibro"
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
            Left            =   2490
            TabIndex        =   49
            Top             =   390
            Width           =   4020
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Libro"
            Height          =   195
            Index           =   3
            Left            =   195
            TabIndex        =   48
            Top             =   480
            Width           =   345
         End
         Begin VB.Label lbl_cb_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb_cod(1)"
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
            Index           =   1
            Left            =   4680
            TabIndex        =   45
            Top             =   735
            Visible         =   0   'False
            Width           =   1230
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
            Height          =   300
            Index           =   1
            Left            =   2490
            TabIndex        =   44
            Top             =   735
            Width           =   4020
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "Sub Libro"
            Height          =   195
            Index           =   1
            Left            =   195
            TabIndex        =   43
            Top             =   800
            Width           =   675
         End
         Begin VB.Line Line2 
            BorderWidth     =   5
            Index           =   0
            X1              =   2550
            X2              =   2535
            Y1              =   2130
            Y2              =   2130
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Glosa"
            Height          =   195
            Index           =   4
            Left            =   195
            TabIndex        =   38
            Top             =   2400
            Width           =   405
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Doc."
            Height          =   195
            Index           =   6
            Left            =   195
            TabIndex        =   37
            Top             =   1440
            Width           =   840
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento"
            Height          =   195
            Index           =   1
            Left            =   195
            TabIndex        =   36
            Top             =   1760
            Width           =   1185
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Nº Documento"
            Height          =   195
            Index           =   2
            Left            =   195
            TabIndex        =   35
            Top             =   2080
            Width           =   1050
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   0
            Left            =   195
            TabIndex        =   32
            Top             =   1120
            Width           =   585
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
            Left            =   2490
            TabIndex        =   31
            Top             =   1050
            Width           =   2055
         End
         Begin VB.Label LblIdCli 
            AutoSize        =   -1  'True
            Caption         =   "LblIdCli"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   8745
            TabIndex        =   27
            Top             =   1335
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Index           =   5
            Left            =   6990
            TabIndex        =   25
            Top             =   1335
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label LblDocumento 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDocumento"
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
            Left            =   2505
            TabIndex        =   24
            Top             =   1680
            Width           =   4020
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Asiento"
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
            Left            =   90
            TabIndex        =   22
            Top             =   30
            Width           =   11610
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
            Top             =   405
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11218
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
            Columns(1).Caption=   "Num.Reg."
            Columns(1).DataField=   "registro"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "T.D."
            Columns(2).DataField=   "destipdoc"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Nº Documento"
            Columns(3).DataField=   "numedoc"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "M"
            Columns(4).DataField=   "simbolo"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Fecha"
            Columns(5).DataField=   "fchdoc"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Sub Libro"
            Columns(6).DataField=   "sublibdesc"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Glosa"
            Columns(7).DataField=   "glosa"
            Columns(7).NumberFormat=   "0.00"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Debe"
            Columns(8).DataField=   "totdeb1"
            Columns(8).NumberFormat=   "0.00"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Haber"
            Columns(9).DataField=   "tothab1"
            Columns(9).NumberFormat=   "0.00"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   10
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=10"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1640"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1561"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=820"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=741"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=714"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=635"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=1508"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1429"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=2514"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=2434"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=5900"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=5821"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=512"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(50)=   "Column(8).Width=1879"
            Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=1799"
            Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=514"
            Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(56)=   "Column(9).Width=1958"
            Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=1879"
            Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=514"
            Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
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
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.fgcolor=&HFFFFFF&"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=74,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=71,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=72,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=73,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=66,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=70,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=54,.parent=13,.alignment=0"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=43,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=44,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=45,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=58,.parent=13,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=55,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=56,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=57,.parent=17"
            _StyleDefs(76)  =   "Named:id=33:Normal"
            _StyleDefs(77)  =   ":id=33,.parent=0"
            _StyleDefs(78)  =   "Named:id=34:Heading"
            _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(80)  =   ":id=34,.wraptext=-1"
            _StyleDefs(81)  =   "Named:id=35:Footing"
            _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(83)  =   "Named:id=36:Selected"
            _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(85)  =   "Named:id=37:Caption"
            _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(87)  =   "Named:id=38:HighlightRow"
            _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(89)  =   "Named:id=39:EvenRow"
            _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(91)  =   "Named:id=40:OddRow"
            _StyleDefs(92)  =   ":id=40,.parent=33"
            _StyleDefs(93)  =   "Named:id=41:RecordSelector"
            _StyleDefs(94)  =   ":id=41,.parent=34"
            _StyleDefs(95)  =   "Named:id=42:FilterBar"
            _StyleDefs(96)  =   ":id=42,.parent=33"
         End
         Begin VB.Label lblperiodo 
            AutoSize        =   -1  'True
            Caption         =   "lblperiodo"
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
            Index           =   0
            Left            =   9450
            TabIndex        =   28
            Top             =   75
            Width           =   1365
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Asientos Diversos"
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
            TabIndex        =   15
            Top             =   90
            Width           =   11595
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
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
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6915
         Top             =   60
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
               Picture         =   "FrmProvisiones2.frx":0B27
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones2.frx":106B
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones2.frx":13FD
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones2.frx":1581
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones2.frx":19D5
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones2.frx":1AED
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones2.frx":2031
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones2.frx":2575
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones2.frx":2689
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones2.frx":279D
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones2.frx":2BF1
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones2.frx":2D5D
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProvisiones2.frx":32A5
               Key             =   "IMG12"
            EndProperty
         EndProperty
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
         Caption         =   "Eliminar             "
      End
   End
End
Attribute VB_Name = "FrmProvisiones2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim CaracteresNumericos As String
Dim RstFrm As New ADODB.Recordset
Dim Agregando As Boolean
Dim RstTmp As New ADODB.Recordset
Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim xId As Double
Dim xHorIni As Date

Dim mMesActivo As Integer '--indica el mes activo
Dim mIdRegistro& '--identificador del registro
Dim fCierrePeriodo As Boolean '--indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO

Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Sub Bloquea()
    TxtTipDoc.Locked = Not TxtTipDoc.Locked
    TxtNombre.Locked = Not TxtNombre.Locked
    TxtSerDoc.Locked = Not TxtSerDoc.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    'If mMesActivo = 0 Or mMesActivo = 13 Then
    TxtFchEmi.Locked = Not TxtFchEmi.Locked
    'End If
    TxtGlosa.Locked = Not TxtGlosa.Locked
   
    habilitar_Locked txt_cb, Not txt_cb(0).Locked
    TxtTC.Locked = Not TxtTC.Locked
    
    ChkAjusteDifCambio.Enabled = Not ChkAjusteDifCambio.Enabled
    
End Sub

Sub Blanquea()
    'TxtIdLibro.Text = ""
    TxtTipDoc.Text = ""
    TxtNombre.Text = ""
    TxtSerDoc.Text = ""
    TxtNumDoc.Text = ""

    TxtFchEmi.Valor = ""
    TxtGlosa.Text = ""
    

    LblLibro.Caption = ""
    LblDocumento.Caption = ""
    LblIdCli.Caption = ""
    TxtTotDeb.Text = ""
    TxtTotHab.Text = ""
    
    TxtTotDebDol.Text = ""
    TxtTotHabDol.Text = ""
    
    LimpiaText txt_cb, True
    
    LblTipoCambio.Caption = ""
    TxtTC.Text = ""
    lblReg.Caption = ""
    
    ChkAjusteDifCambio = 0
    
End Sub

Private Sub CmdAcepta_Click()
    If TxtIdMod.Text = "" Then
        MsgBox "No ha especificado el modulo al que corresponden los documentos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMod.SetFocus
        Exit Sub
    End If
    
    If OptDebe.Value = False And OptHaber.Value = False Then
        MsgBox "No ha especificado la naturaleza de los documentos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    If NulosN(txt_cb(0).Text) = 1 Then
        If OptDebe.Value = True Then Fg1.TextMatrix(Fg1.Row, 3) = SumaColumna(fg2, 21)
        If OptHaber.Value = True Then Fg1.TextMatrix(Fg1.Row, 4) = SumaColumna(fg2, 21)
        Fg1.TextMatrix(Fg1.Row, 9) = Format(NulosN(Fg1.TextMatrix(Fg1.Row, 3)) / NulosN(TxtTC.Text), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Row, 10) = Format(NulosN(Fg1.TextMatrix(Fg1.Row, 4)) / NulosN(TxtTC.Text), FORMAT_MONTO)
    End If
    If NulosN(txt_cb(0).Text) = 2 Then
        If OptDebe.Value = True Then Fg1.TextMatrix(Fg1.Row, 9) = SumaColumna(fg2, 22)
        If OptHaber.Value = True Then Fg1.TextMatrix(Fg1.Row, 10) = SumaColumna(fg2, 22)
        Fg1.TextMatrix(Fg1.Row, 3) = Format(NulosN(Fg1.TextMatrix(Fg1.Row, 9)) * NulosN(TxtTC.Text), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Row, 4) = Format(NulosN(Fg1.TextMatrix(Fg1.Row, 10)) * NulosN(TxtTC.Text), FORMAT_MONTO)
    End If
    
    TxtTotDeb.Text = Format(SumaColumna(Fg1, 3), FORMAT_MONTO)
    TxtTotHab.Text = Format(SumaColumna(Fg1, 4), FORMAT_MONTO)
    
    TxtTotDebDol.Text = Format(SumaColumna(Fg1, 9), FORMAT_MONTO)
    TxtTotHabDol.Text = Format(SumaColumna(Fg1, 10), FORMAT_MONTO)
    
    TabOne1.Enabled = True
    Toolbar1.Enabled = True
    Frame7.Visible = False
End Sub

Function SumaColumna(xFlex As VSFlexGrid, Columna As Integer) As Double
    Dim A As Integer
    Dim xTotal As Double
    For A = 1 To xFlex.Rows - 1
        xTotal = xTotal + NulosN(xFlex.TextMatrix(A, Columna))
    Next A
    SumaColumna = xTotal
End Function

Private Sub CmdAdd_Click()
    If QueHace = 3 Then Exit Sub
    If Fg1.TextMatrix(Fg1.Rows - 1, 1) = "" Then
        Fg1.Row = Fg1.Rows - 1
        Fg1.Col = 1
        Fg1.SetFocus
        If Fg1.Row < Fg1.FixedRows Or Fg1.Col < Fg1.FixedCols Then Exit Sub
        Fg1_CellButtonClick Fg1.Row, Fg1.Col
        Exit Sub
    End If
    Fg1.Rows = Fg1.Rows + 1
    Fg1.Row = Fg1.Rows - 1:        Fg1.Col = 1
    Fg1.SetFocus
    If Fg1.Row < Fg1.FixedRows Or Fg1.Col < Fg1.FixedCols Then Exit Sub
    Fg1_CellButtonClick Fg1.Row, Fg1.Col
End Sub

Private Sub CmdAdd2_Click()
    Dim xIdDoc As Integer
    fg2.Rows = fg2.Rows + 1
    RstTmp.Filter = adFilterNone
    RstTmp.Filter = "idprovi = " & xId & " AND idcuent = " & NulosN(Fg1.TextMatrix(Fg1.Row, 5)) & ""
    
    If RstTmp.RecordCount <> 0 Then
        RstTmp.Sort = "iddoc"
        RstTmp.MoveLast
        xIdDoc = RstTmp("iddoc") + 1
    Else
        xIdDoc = 1
    End If
    RstTmp.AddNew
    RstTmp("idprovi") = xId
    RstTmp("idcuent") = Fg1.TextMatrix(Fg1.Row, 5)
    RstTmp("iddoc") = xIdDoc
    RstTmp("nuevo") = 1
    
    Agregando = True
    fg2.TextMatrix(fg2.Rows - 1, 13) = xIdDoc
    fg2.TextMatrix(fg2.Rows - 1, 14) = Fg1.TextMatrix(Fg1.Row, 5)
    fg2.TextMatrix(fg2.Rows - 1, 15) = xId

    Agregando = False
End Sub

Private Sub CmdBusLib_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(2, 4) As String
    
    xCampos2(0, 0) = "Documento":    xCampos2(0, 1) = "descripcion":    xCampos2(0, 2) = "5000":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Codigo":       xCampos2(1, 1) = "id":             xCampos2(1, 2) = "1000":         xCampos2(1, 3) = "N"

    xform.SqlCad = "SELECT * FROM mae_libros"
    xform.Titulo = "Buscando Libros Contables"
        
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos2)
    
    If xRs.State = 1 Then
        TxtIdLibro.Text = NulosN(xRs("id"))
        LblLibro.Caption = NulosC(xRs("descripcion"))
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMod_Click()
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(2, 4) As String
    
    xCampos2(0, 0) = "Modulo":    xCampos2(0, 1) = "descripcion":    xCampos2(0, 2) = "5000":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Codigo":    xCampos2(1, 1) = "id":             xCampos2(1, 2) = "1000":         xCampos2(1, 3) = "N"

    xform.SqlCad = "SELECT * FROM tes_modulos"
    xform.Titulo = "Buscando Modulos"
        
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos2)
    
    If xRs.State = 1 Then
        TxtIdMod.Text = xRs("id")
        LblDescModulo.Caption = xRs("descripcion")
        Fg1.TextMatrix(Fg1.Row, 6) = xRs("id")
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipDoc_Click()
    If QueHace = 3 Then Exit Sub
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(3, 4) As String
    
    xCampos2(0, 0) = "Documento":    xCampos2(0, 1) = "descripcion":    xCampos2(0, 2) = "5000":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Sigla":        xCampos2(1, 1) = "abrev":          xCampos2(1, 2) = "1000":         xCampos2(1, 3) = "N"
    xCampos2(2, 0) = "Codigo":       xCampos2(2, 1) = "id":             xCampos2(2, 2) = "1000":         xCampos2(2, 3) = "N"

    xform.SqlCad = "SELECT * FROM mae_documento"
    xform.Titulo = "Buscando Documentos"
        
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos2)
    
    If xRs.State = 1 Then
        TxtTipDoc.Text = xRs("id")
        LblDocumento.Caption = xRs("descripcion")
        TxtSerDoc.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdDel_Click()
    If QueHace = 3 Then Exit Sub
    If Fg1.Rows = 1 Then
        MsgBox "No hay cuentas para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    If Fg1.Row < 1 Then Exit Sub
    If MsgBox("Seguro desea eliminar el registro" + vbCr + "N° Cuenta: " + Fg1.TextMatrix(Fg1.Row, 1) + vbCr + "Descripción: " + Fg1.TextMatrix(Fg1.Row, 2), vbQuestion + vbYesNo + vbDefaultButton1, xTitulo) = vbNo Then Exit Sub
    Fg1.RemoveItem Fg1.Row
    HallarTotal
End Sub

Private Sub CmdDel2_Click()
    If NulosN(fg2.TextMatrix(fg2.Row, 19)) = 1 Then 'preguntamos si es editable
        MsgBox "No se puede eliminar este documento, tiene movimientos en caja y bancos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    Else
        RstTmp.MoveFirst
        RstTmp.Filter = adFilterNone
        
        RstTmp.Filter = adFilterNone
        RstTmp.Filter = "idprovi = " & xId & " AND idcuent = " & NulosN(Fg1.TextMatrix(Fg1.Row, 5)) & " AND iddoc = " & NulosN(fg2.TextMatrix(fg2.Row, 13)) & ""
        If RstTmp.RecordCount <> 0 Then
            RstTmp.Delete
        End If
        fg2.RemoveItem fg2.Row
    
    End If
End Sub

Private Sub CmdDelTodo_Click()
    Dim Rst As New ADODB.Recordset
    Dim A, B, Borrados, NoBorrados As Integer
    
    Borrados = 0
    NoBorrados = 0
    
    If fg2.TextMatrix(fg2.Row, 14) = 1 Then 'compras
        For B = 1 To fg2.Rows - 1
            RST_Busq Rst, "SELECT tes_cajadestinodet.iddoc, tes_cajadestinodet.idmod, tes_cajadestinodet.acuenta From tes_cajadestinodet " _
                & " WHERE (((tes_cajadestinodet.iddoc)=" & NulosN(fg2.TextMatrix(B, 12)) & ") AND ((tes_cajadestinodet.idmod)=1) AND ((tes_cajadestinodet.acuenta)<>0))", xCon
    
            If Rst.RecordCount <> 0 Then
                'MsgBox "No se puede eliminar este documento, tiene movimientos en Caja y Bancos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                NoBorrados = NoBorrados + 1
                Set Rst = Nothing
            Else
                RstTmp.Filter = adFilterNone
                RstTmp.MoveFirst
                RstTmp.Filter = "iddoc = " & NulosN(fg2.TextMatrix(B, 12)) & " AND idprovi = " & RstFrm("id") & " AND idcuent = " & Fg1.TextMatrix(Fg1.Row, 5) & ""
                If RstTmp.RecordCount <> 0 Then
                    RstTmp.MoveFirst
                    For A = 1 To RstTmp.RecordCount
                        RstTmp.Delete
                        RstTmp.MoveNext
                        If RstTmp.EOF = True Then Exit For
                    Next A
                End If
                Borrados = Borrados + 1
                fg2.RemoveItem B
                
                B = B - 1
            End If
            If B = fg2.Rows - 1 Then Exit For
        Next B
        MsgBox "Se han Borrado " + Trim(Str(Borrados)) + " Documentos de Compra" & Chr(13) & _
               "No se Borraron " + Trim(Str(NoBorrados)) + " por tener movimientos en caja y bancos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    If fg2.TextMatrix(fg2.Row, 14) = 2 Then 'ventas
        
    End If
    
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstFrm
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

Private Sub Dg1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If
End Sub

Sub MostrarDetalleProvicion(IdProvicion As Double, IdCuenta As Integer)
    fg2.Rows = 1
    Dim A As Integer
    Dim xTotal1, xTotal2, xTotal3, xTotal4, xTotal5 As Double
    
    RstTmp.Filter = adFilterNone
    RstTmp.Filter = "idprovi = " & IdProvicion & " and idcuent  = " & IdCuenta & ""
    
    If RstTmp.RecordCount <> 0 Then
        RstTmp.MoveFirst
        Agregando = True
        For A = 1 To RstTmp.RecordCount
            fg2.Rows = fg2.Rows + 1
            
            fg2.TextMatrix(A, 1) = Format(RstTmp("fchdoc"), "dd/mm/yy")
            fg2.TextMatrix(A, 2) = NulosC(RstTmp("desdoc"))
            fg2.TextMatrix(A, 3) = RstTmp("numdoc")
            fg2.TextMatrix(A, 4) = RstTmp("nommon")
            fg2.TextMatrix(A, 5) = RstTmp("provee")
            fg2.TextMatrix(A, 6) = RstTmp("nomcon")
            fg2.TextMatrix(A, 7) = Format(RstTmp("fchven"), "dd/mm/yy")
            fg2.TextMatrix(A, 8) = Format(RstTmp("impbru"), FORMAT_MONTO)
            fg2.TextMatrix(A, 9) = Format(RstTmp("impigv"), FORMAT_MONTO)
            fg2.TextMatrix(A, 10) = Format(RstTmp("impisc"), FORMAT_MONTO)
            fg2.TextMatrix(A, 11) = Format(RstTmp("imptot"), FORMAT_MONTO)
            fg2.TextMatrix(A, 12) = Format(RstTmp("impsal"), FORMAT_MONTO)
            
            fg2.TextMatrix(A, 13) = RstTmp("iddoc")
            fg2.TextMatrix(A, 15) = RstTmp("idprovi")
            fg2.TextMatrix(A, 16) = RstTmp("idmon")
            fg2.TextMatrix(A, 17) = RstTmp("idpro")
            fg2.TextMatrix(A, 18) = RstTmp("idcon")
            fg2.TextMatrix(A, 19) = RstTmp("edita")
            fg2.TextMatrix(A, 20) = RstTmp("tipdoc")
            
            ConvertirMoneda A
            
            If RstTmp("edita") = 0 Then
                '0 =  no se edita la fila
                GRID_COLOR_FONDO fg2, CLng(A), 1, CLng(A), 18, &HFFFFFF
            Else
                '1 = se edita la fila
                GRID_COLOR_FONDO fg2, CLng(A), 1, CLng(A), 18, &HFFFF&
            End If
            
            RstTmp.MoveNext
            If RstTmp.EOF = True Then Exit For
        Next A
        Agregando = False
    End If
    
    If NulosN(Fg1.TextMatrix(Fg1.Row, 3)) <> 0 Then OptDebe.Value = True
    If NulosN(Fg1.TextMatrix(Fg1.Row, 4)) <> 0 Then OptHaber.Value = True
    
    TxtIdMod.Text = NulosN(Fg1.TextMatrix(Fg1.Row, 6))
    LblDescModulo.Caption = Busca_Codigo(NulosN(Fg1.TextMatrix(Fg1.Row, 6)), "id", "descripcion", "tes_modulos", "N", xCon)
    
    If QueHace = 3 Then
        CmdAdd2.Enabled = False
        CmdDel2.Enabled = False
        CmdDelTodo.Enabled = False
        Command2.Enabled = False
        Command3.Enabled = False
    Else
        CmdAdd2.Enabled = True
        CmdDel2.Enabled = True
        CmdDelTodo.Enabled = True
        Command2.Enabled = True
        Command3.Enabled = True
    End If
    TabOne1.Enabled = False
    Toolbar1.Enabled = False
    Frame7.Left = 30
    Frame7.Top = 1695
    Frame7.Visible = True
    
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstFrm("id")), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    'If QueHace = 3 Then Exit Sub
  
    If Col = 3 Or Col = 4 Or Col = 9 Or Col = 10 Then
        If QueHace = 3 Then xId = RstFrm("id")
        MostrarDetalleProvicion xId, Fg1.TextMatrix(Fg1.Row, 5)
    End If

    If Col = 1 Then
        Dim Rst As New ADODB.Recordset
        Dim xRs As New ADODB.Recordset
        Dim nSQL As String
        Dim nSQLLike As String
        Dim nSQLIdCta As String
          
        Dim xCampos(3, 4) As String
        
        xCampos(0, 0) = "Nro Cta":         xCampos(0, 1) = "cuenta":             xCampos(0, 2) = "1500":    xCampos(0, 3) = "C"
        xCampos(1, 0) = "Nombre Cuenta":    xCampos(1, 1) = "descripcion":       xCampos(1, 2) = "5000":    xCampos(1, 3) = "C"
        xCampos(2, 0) = "Divisionaria":   xCampos(2, 1) = "xtipo":               xCampos(2, 2) = "1200":    xCampos(2, 3) = "C"
        
        If NulosC(Fg1.TextMatrix(Fg1.Row, 1)) <> "" Then
            Fg1.TextMatrix(Fg1.Row, 1) = Replace(Fg1.TextMatrix(Fg1.Row, 1), "'", "")
            Fg1.TextMatrix(Fg1.Row, 1) = Replace(Fg1.TextMatrix(Fg1.Row, 1), "*", "")
            Fg1.TextMatrix(Fg1.Row, 1) = Replace(Fg1.TextMatrix(Fg1.Row, 1), "LIKE", "")
            
            nSQLLike = " and con_planctas.cuenta like '" + Trim(Fg1.TextMatrix(Fg1.Row, 1)) + "%' "
            
        End If
           
        nSQLIdCta = GRID_GENERAR_SQL_ID(Fg1, 5, " and con_planctas.id", " NOT IN ", True)
        
        nSQL = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id, con_planctas.documentar, con_planctas.idmodulo,con_planctas.tipo,iif(con_planctas.tipo=0,'Si','No') as xtipo " _
            + vbCr + " From con_planctas where con_planctas.id <>0 " + nSQLIdCta + nSQLLike + vbCr + "  ORDER BY con_planctas.cuenta"
                
        CARGAR_DLL_EPSBUSCAR xCon, Rst, nSQL, xCampos(), "Buscando Cuentas Contables", "cuenta", "cuenta", Principio
        
        If Rst.State = 0 Then GoTo SALIR
        If Rst.RecordCount = 0 Then GoTo SALIR
           
'        RST_Busq xRs, "SELECT id, cuenta FROM con_planctas WHERE (((id)<>" + Trim(Rst("id")) + ") AND ((cuenta) Like '" + Trim(Rst("cuenta")) + "%'));", xCon
'        If xRs.EOF = False Or xRs.BOF = False Or xRs.RecordCount <> 0 Then
'            MsgBox "Cuenta no válida" + vbCr + "Seleccione una divisionaria", vbExclamation, xTitulo
'            GoTo SALIR
'            Exit Sub
'        End If

        '--validar que solo se agregue cuentas divisionarias
        If NulosN(Rst("tipo")) = 1 Then
            MsgBox "Cuenta no válida" + vbCr + "Seleccione una divisionaria", vbExclamation, xTitulo
            GoTo SALIR
            Exit Sub
        End If
        
        Agregando = True
    
        If GRID_BUSCAR_VALOR(Fg1, 1, Trim(Rst("cuenta")), False, , Row) <> "-1" Then
            MsgBox "La Cuenta " + Trim(Rst("cuenta")) + " ya esta en la Lista" + vbCr + "Seleccione otra", vbExclamation, xTitulo
            GoTo SALIR
        End If
        Fg1.TextMatrix(Fg1.Row, 1) = NulosC(Rst("cuenta"))
        Fg1.TextMatrix(Fg1.Row, 2) = NulosC(Rst("descripcion"))
        Fg1.TextMatrix(Fg1.Row, 5) = NulosN(Rst("id"))
        
        Fg1.TextMatrix(Fg1.Row, 8) = NulosN(Rst("documentar"))
        Fg1.TextMatrix(Fg1.Row, 6) = NulosN(Rst("idmodulo"))
        
        Fg1.TextMatrix(Fg1.Row, 3) = "0.00"
        Fg1.TextMatrix(Fg1.Row, 4) = "0.00"
        Fg1.TextMatrix(Fg1.Row, 9) = "0.00"
        Fg1.TextMatrix(Fg1.Row, 10) = "0.00"
        
        
        Set Rst = Nothing
        Set xRs = Nothing
    End If
    
SALIR:
    
    Agregando = False
    Exit Sub
error:
    'Resume
    Set Rst = Nothing
    Set xRs = Nothing
    Agregando = False
    SHOW_ERROR Me.Name, "Fg1_CellButtonClick"
End Sub

Sub HallarTotal()
    TxtTotDeb.Text = Format(GRID_SUMAR_COL(Fg1, 3), FORMAT_MONTO)
    TxtTotHab.Text = Format(GRID_SUMAR_COL(Fg1, 4), FORMAT_MONTO)
    
    TxtTotDebDol.Text = Format(GRID_SUMAR_COL(Fg1, 9), FORMAT_MONTO)
    TxtTotHabDol.Text = Format(GRID_SUMAR_COL(Fg1, 10), FORMAT_MONTO)
    
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    If Fg1.TextMatrix(Row, Col) = "" Then
        Fg1.TextMatrix(Row, 2) = ""
        Fg1.TextMatrix(Row, 5) = ""
        Exit Sub
    End If
    
    If Col = 1 Then
        If GRID_BUSCAR_VALOR(Fg1, 1, Trim(Fg1.TextMatrix(Row, Col)), False, -1, Row) <> "-1" Then
            MsgBox "El Num. Cuenta Contable ya existe" + vbCr + "Ingrese otro Num. Cuenta Contable", vbExclamation, xTitulo
            Fg1.TextMatrix(Row, 1) = ""
            Fg1.TextMatrix(Row, 5) = ""
            Exit Sub
        End If
        
        Dim Rst As New ADODB.Recordset
        RST_Busq Rst, "SELECT * FROM con_planctas WHERE cuenta = '" & NulosC(Fg1.TextMatrix(Row, 1)) & "'", xCon
        If Rst.RecordCount = 1 Then
            Fg1.TextMatrix(Row, 2) = NulosC(Rst("descripcion"))
            Fg1.TextMatrix(Row, 5) = NulosN(Rst("id"))
        Else
            Fg1.TextMatrix(Row, 2) = ""
            Fg1.TextMatrix(Row, 5) = ""
        End If
        Set Rst = Nothing
    End If
    
    If Col = 3 Or Col = 4 Or Col = 9 Or Col = 10 Then
        If IsNumeric(Fg1.TextMatrix(Row, Col)) = False Then
            MsgBox "El valor ingresado no es numérico", vbCritical, xTitulo
            Fg1.TextMatrix(Row, Col) = ""
        Else
            If Col = 3 And NulosN(Fg1.TextMatrix(Row, 4)) > 0 Then
                Fg1.TextMatrix(Row, 4) = 0
            ElseIf Col = 4 And NulosN(Fg1.TextMatrix(Row, 3)) > 0 Then
                Fg1.TextMatrix(Row, 3) = 0
            ElseIf Col = 9 And NulosN(Fg1.TextMatrix(Row, 10)) > 0 Then
                Fg1.TextMatrix(Row, 10) = 0
            ElseIf Col = 10 And NulosN(Fg1.TextMatrix(Row, 9)) > 0 Then
                Fg1.TextMatrix(Row, 9) = 0
            End If
        End If
        
        If NulosN(Fg1.TextMatrix(Fg1.Row, 6)) = 0 Then
            If NulosN(TxtTC.Text) <> 0 Then
                If Col = 3 And ChkAjusteDifCambio.Value = 0 Then Fg1.TextMatrix(Fg1.Row, 9) = NulosN(Fg1.TextMatrix(Fg1.Row, 3)) / NulosN(TxtTC.Text)
                If Col = 4 And ChkAjusteDifCambio.Value = 0 Then Fg1.TextMatrix(Fg1.Row, 10) = NulosN(Fg1.TextMatrix(Fg1.Row, 4)) / NulosN(TxtTC.Text)
            End If
            If Col = 9 And ChkAjusteDifCambio.Value = 0 Then Fg1.TextMatrix(Fg1.Row, 3) = NulosN(Fg1.TextMatrix(Fg1.Row, 9)) * NulosN(TxtTC.Text)
            If Col = 10 And ChkAjusteDifCambio.Value = 0 Then Fg1.TextMatrix(Fg1.Row, 4) = NulosN(Fg1.TextMatrix(Fg1.Row, 10)) * NulosN(TxtTC.Text)
            
        End If
        
        Fg1.TextMatrix(Fg1.Row, 3) = Format(Fg1.TextMatrix(Fg1.Row, 3), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Row, 4) = Format(Fg1.TextMatrix(Fg1.Row, 4), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Row, 9) = Format(Fg1.TextMatrix(Fg1.Row, 9), FORMAT_MONTO)
        Fg1.TextMatrix(Fg1.Row, 10) = Format(Fg1.TextMatrix(Fg1.Row, 10), FORMAT_MONTO)
        
    End If
    
    TxtTotDeb.Text = Format(SumaColumna(Fg1, 3), FORMAT_MONTO)
    TxtTotHab.Text = Format(SumaColumna(Fg1, 4), FORMAT_MONTO)
    TxtTotDebDol.Text = Format(SumaColumna(Fg1, 9), FORMAT_MONTO)
    TxtTotHabDol.Text = Format(SumaColumna(Fg1, 10), FORMAT_MONTO)
End Sub

Private Sub fg1_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Fg1_CellButtonClick Row, Col
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        
        If Fg1.Col = 3 Or Fg1.Col = 4 Or Fg1.Col = 9 Or Fg1.Col = 10 Then
            If NulosN(Fg1.TextMatrix(Fg1.Row, 6)) = 0 Then
                Fg1.ColComboList(3) = ""
                Fg1.ColComboList(4) = ""
                Fg1.ColComboList(9) = ""
                Fg1.ColComboList(10) = ""
            Else
                If NulosN(txt_cb(0).Text) = 1 Then
                    Fg1.ColComboList(3) = "|..."
                    Fg1.ColComboList(4) = "|..."
                    Fg1.ColComboList(9) = ""
                    Fg1.ColComboList(10) = ""
                End If
                If NulosN(txt_cb(0).Text) = 2 Then
                    Fg1.ColComboList(9) = "|..."
                    Fg1.ColComboList(10) = "|..."
                    Fg1.ColComboList(3) = ""
                    Fg1.ColComboList(4) = ""
                End If
            
                Fg1.Editable = flexEDKbdMouse
            End If
        End If
        Exit Sub
    End If
    
    If Fg1.Col = 1 Or Fg1.Col = 3 Or Fg1.Col = 4 Or Fg1.Col = 9 Or Fg1.Col = 10 Then
        If Fg1.Col = 1 Then Fg1.Editable = flexEDKbdMouse
            
        If Fg1.Col = 3 Or Fg1.Col = 4 Or Fg1.Col = 9 Or Fg1.Col = 10 Then
            Fg1.Editable = flexEDKbdMouse
            If NulosN(txt_cb(0).Text) = 1 Then
                If NulosN(Fg1.TextMatrix(Fg1.Row, 6)) = 0 Then
                    Fg1.ColComboList(3) = ""
                    Fg1.ColComboList(4) = ""
                Else
                    Fg1.ColComboList(3) = "|..."
                    Fg1.ColComboList(4) = "|..."
                End If
                
                Fg1.ColComboList(9) = ""
                Fg1.ColComboList(10) = ""
            End If
            
            If NulosN(txt_cb(0).Text) = 2 Then
                If NulosN(Fg1.TextMatrix(Fg1.Row, 6)) = 0 Then
                    Fg1.ColComboList(9) = ""
                    Fg1.ColComboList(10) = ""
                Else
                    Fg1.ColComboList(9) = "|..."
                    Fg1.ColComboList(10) = "|..."
                End If
                
                Fg1.ColComboList(3) = ""
                Fg1.ColComboList(4) = ""
            End If
        End If
        
        If NulosN(txt_cb(0).Text) = 1 Then
            If Fg1.Col = 9 Or Fg1.Col = 10 Then
                Fg1.Editable = flexEDNone
            End If
            If Fg1.Col = 3 Or Fg1.Col = 4 Then
                Fg1.Editable = flexEDKbdMouse
            End If
        End If
        
        If NulosN(txt_cb(0).Text) = 2 Then
            If Fg1.Col = 3 Or Fg1.Col = 4 Then
                Fg1.Editable = flexEDNone
            End If
            If Fg1.Col = 9 Or Fg1.Col = 10 Then
                Fg1.Editable = flexEDKbdMouse
            End If
        End If
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If QueHace = 3 Or Row < 1 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    Select Case Col
        Case 1
            
        Case 3, 4, 9, 10
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    
    If KeyCode = 45 Then
        CmdAdd_Click
    End If
    
    If KeyCode = 46 Then
        CmdDel_Click
    End If
    
    If KeyCode = 121 Then
        If Fg1.Row < Fg1.FixedRows Or Fg1.Col < Fg1.FixedCols Then Exit Sub
        Fg1_CellButtonClick Fg1.Row, Fg1.Col
    End If


    
    
    If KeyCode = 122 Then
        If NulosN(txt_cb(0).Text) = 1 Then
            If Fg1.Col = 9 Or Fg1.Col = 10 Then
                Fg1.Editable = flexEDKbdMouse
                SendKeys "{ENTER}"
            End If
        End If
        
        If NulosN(txt_cb(0).Text) = 2 Then
            If Fg1.Col = 3 Or Fg1.Col = 4 Then
                Fg1.Editable = flexEDKbdMouse
                SendKeys "{ENTER}"
            End If
        End If
    End If
    
    
End Sub

Private Sub Fg1_LostFocus()
    HallarTotal
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace <> 3 Then
            PopupMenu menu1
        End If
    End If
End Sub

Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    If Col = 2 Then   'Moneda
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "id":            xCampos(0, 2) = "1200":         xCampos(0, 3) = "N"
        xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
        
        'filtramos por tipo de movimiento  = 1 (Ingreso)
        xform.SqlCad = "SELECT * FROM  mae_documento ORDER BY descripcion"
    
        xform.Titulo = "Buscando Documento"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "id"
        xform.CampoBusca = "id"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 1 Then
            fg2.TextMatrix(fg2.Row, 20) = xRs("id")
            fg2.TextMatrix(fg2.Row, 2) = xRs("abrev")
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If
     
    If Col = 4 Then   'Moneda
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "id":            xCampos(0, 2) = "1200":         xCampos(0, 3) = "N"
        xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
        
        'filtramos por tipo de movimiento  = 1 (Ingreso)
        xform.SqlCad = "SELECT * FROM  mae_moneda ORDER BY descripcion"
    
        xform.Titulo = "Buscando Moneda"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "id"
        xform.CampoBusca = "id"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 1 Then
            fg2.TextMatrix(fg2.Row, 16) = xRs("id")
            fg2.TextMatrix(fg2.Row, 4) = xRs("simbolo")
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If
    
    If Col = 5 Then  ' Proveedor
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        Dim xCampos2(3, 4) As String
        
        If Fg1.TextMatrix(Fg1.Row, 6) = 1 Then
            xCampos2(0, 0) = "Proveedor":     xCampos2(0, 1) = "nombre":        xCampos2(0, 2) = "4000":    xCampos2(0, 3) = "C"
            xform.SqlCad = "SELECT mae_prov.id, mae_prov.numruc, mae_prov.nombre FROM mae_prov"
            xform.Titulo = "Buscando Proveedor"
        End If
        If Fg1.TextMatrix(Fg1.Row, 6) = 2 Then
            xCampos2(0, 0) = "Cliente":     xCampos2(0, 1) = "nombre":        xCampos2(0, 2) = "4000":    xCampos2(0, 3) = "C"
            xform.SqlCad = "SELECT mae_cliente.id, mae_cliente.numruc, mae_cliente.nombre FROM mae_cliente"
            xform.Titulo = "Buscando Cliente"
        End If
        
        xCampos2(1, 0) = "Nº R.U.C.":     xCampos2(1, 1) = "numruc":        xCampos2(1, 2) = "1000":    xCampos2(1, 3) = "C"
        xCampos2(2, 0) = "Codigo":        xCampos2(2, 1) = "id":            xCampos2(2, 2) = "1000":    xCampos2(2, 3) = "N"
        
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "nombre"
        xform.CampoBusca = "nombre"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos2)
        If xRs.State = 1 Then
            fg2.TextMatrix(fg2.Row, 17) = xRs("id")
            fg2.TextMatrix(fg2.Row, 5) = xRs("nombre")
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If
    
    If Col = 6 Then  'condicion de pago
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "id":            xCampos(0, 2) = "1200":         xCampos(0, 3) = "N"
        xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
        
        'filtramos por tipo de movimiento  = 1 (Ingreso)
        xform.SqlCad = "SELECT * FROM  mae_condpago ORDER BY descripcion"
    
        xform.Titulo = "Buscando Condicion de Pago"
        xform.FormaBusca = Principio
        xform.Criterio = ""
        xform.Ordenado = "id"
        xform.CampoBusca = "id"
        Set xform.Coneccion = xCon
        Set xRs = xform.BuscarReg(xCampos)
        If xRs.State = 1 Then
            fg2.TextMatrix(fg2.Row, 18) = xRs("id")
            fg2.TextMatrix(fg2.Row, 6) = xRs("abrev")
        End If
        Set xform = Nothing
        Set xRs = Nothing
    End If
End Sub

Private Sub Fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    RstTmp.Filter = adFilterNone
    RstTmp.MoveFirst
        
    RstTmp.Filter = "iddoc = " & NulosN(fg2.TextMatrix(fg2.Row, 13)) & " AND idprovi = " & xId & "" _
        & " AND idcuent = " & NulosN(Fg1.TextMatrix(Fg1.Row, 5)) & ""
    
   If RstTmp.RecordCount <> 0 Then
        If Col = 1 Then RstTmp("fchdoc") = fg2.TextMatrix(fg2.Row, 1) 'fecha del documento
        If Col = 2 Then
            RstTmp("tipdoc") = NulosN(fg2.TextMatrix(fg2.Row, 20))  'codifo del documento
            RstTmp("desdoc") = NulosC(fg2.TextMatrix(fg2.Row, 2)) 'descripcion del documento
        End If
        
        If Col = 3 Then RstTmp("numdoc") = fg2.TextMatrix(fg2.Row, 3) 'numero del documento
        
        If Col = 4 Then
            RstTmp("idmon") = NulosN(fg2.TextMatrix(fg2.Row, 16)) 'codigo de la moneda
            RstTmp("nommon") = fg2.TextMatrix(fg2.Row, 4) 'nombre de la moneda
        End If
        
        If Col = 5 Then
            RstTmp("idpro") = NulosN(fg2.TextMatrix(fg2.Row, 17)) 'id del proveedor
            RstTmp("provee") = fg2.TextMatrix(fg2.Row, 5) 'nombre del proveedor
        End If
        
        If Col = 6 Then
            RstTmp("idcon") = NulosN(fg2.TextMatrix(fg2.Row, 18))
            RstTmp("nomcon") = fg2.TextMatrix(fg2.Row, 6)
        End If
        
        If Col = 7 Then RstTmp("fchven") = fg2.TextMatrix(fg2.Row, 7)
        
        If Col = 8 Then RstTmp("impbru") = fg2.TextMatrix(fg2.Row, 8)
        If Col = 9 Then RstTmp("impigv") = fg2.TextMatrix(fg2.Row, 9)
        If Col = 10 Then RstTmp("impisc") = fg2.TextMatrix(fg2.Row, 10)
        If Col = 11 Then RstTmp("imptot") = fg2.TextMatrix(fg2.Row, 11): fg2.TextMatrix(fg2.Row, 11) = Format(fg2.TextMatrix(fg2.Row, 11), FORMAT_MONTO)
        If Col = 12 Then RstTmp("impsal") = fg2.TextMatrix(fg2.Row, 12): fg2.TextMatrix(fg2.Row, 12) = Format(fg2.TextMatrix(fg2.Row, 12), FORMAT_MONTO)
        
        ConvertirMoneda fg2.Row
    End If
End Sub

Sub ConvertirMoneda(xFila As Integer)
    If NulosN(fg2.TextMatrix(xFila, 12)) <> 0 And NulosN(TxtTC.Text) <> 0 Then
        If NulosN(fg2.TextMatrix(xFila, 16)) = 1 Then
            fg2.TextMatrix(xFila, 21) = Format(fg2.TextMatrix(xFila, 12), FORMAT_MONTO)
            fg2.TextMatrix(xFila, 22) = NulosN(fg2.TextMatrix(xFila, 12)) / NulosN(TxtTC.Text)
            fg2.TextMatrix(xFila, 22) = Format(fg2.TextMatrix(xFila, 22), FORMAT_MONTO)
        End If
        If NulosN(fg2.TextMatrix(xFila, 16)) = 2 Then
            fg2.TextMatrix(xFila, 21) = NulosN(fg2.TextMatrix(xFila, 12)) * NulosN(TxtTC.Text)
            fg2.TextMatrix(xFila, 21) = Format(fg2.TextMatrix(xFila, 21), FORMAT_MONTO)
            fg2.TextMatrix(xFila, 22) = Format(fg2.TextMatrix(xFila, 12), FORMAT_MONTO)
        End If
    End If
End Sub
Private Sub Fg2_EnterCell()
    If QueHace = 3 Then
        fg2.Editable = flexEDNone
    Else
        fg2.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub Fg2_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 45 Then
        CmdAdd2_Click
    End If
    
    If KeyCode = 46 Then
        CmdDel2_Click
    End If
End Sub

Private Sub Form_Activate()

    If SeEjecuto = False Then
       
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        mMesActivo = xMes
        
        OpcionesPeriodo
        
    End If
End Sub

Sub Nuevo()
    
    If PuedeAgregarRegistro("PROVISIONES", xCon) = False Then
        MsgBox "Esta utilizando una versión de prueba del maravilloso sistema SEVEN Soft, si desea la version comercial contactese con el " & Chr(13) _
            & " extraordinario programador Enrique Pollongo a eps_76@hotmail.com y solicite un número de licencia para esta PC", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    QueHace = 1
    ActivaTool
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Agregando Asiento"
    Blanquea
    Bloquea
    
    Fg1.Rows = 1
        
    Fg1.SelectionMode = flexSelectionFree
    Fg1.Editable = flexEDKbdMouse
    Fg1.Rows = Fg1.Rows + 1
    
    PreparaRST_Tmp
    
    TxtIdLibro.Text = 3
    TxtIdLibro_Validate False
    If mMesActivo = 0 Then
        TxtFchEmi.Valor = CDate("31/" + Format(12, "00") + "/" + Trim(Str(AnoTra - 1)))
    End If
    xHorIni = Time
    
    txt_cb(1).SetFocus
    
End Sub

Sub Modificar()
    ActivaTool
    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Modificando Asiento"
    Bloquea
    
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    Fg1.Editable = flexEDKbdMouse
    
    fg2.ColComboList(3) = "|..."
    fg2.ColComboList(4) = "|..."
    fg2.ColComboList(5) = "|..."
    QueHace = 2
    xId = RstFrm("id")
    If mMesActivo = 0 Then
        TxtFchEmi.Valor = CDate("31/" + Format(12, "00") + "/" + Trim(Str(AnoTra - 1)))
    End If
    xHorIni = Time
    txt_cb(1).SetFocus
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    
    If RstFrm.RecordCount = 0 Or RstFrm.EOF = True Or RstFrm.BOF = True Then
        MsgBox "No hay registro para eliminar", vbExclamation, xTitulo
        TabOne1.CurrTab = 0
        Exit Sub
    End If
    
    Rpta = MsgBox("¿Esta seguro de eliminar el asiento seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        Dim Rst As New ADODB.Recordset
        Dim Rst2 As New ADODB.Recordset
        Dim A, B As Integer
        RST_Busq Rst, "SELECT * FROM tes_modulos ORDER BY id", xCon
        
        On Error GoTo LaCague
        xCon.BeginTrans
        
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            RST_Busq Rst2, "SELECT con_provicionesdetdoc.*  From con_provicionesdetdoc " _
                & " WHERE (((con_provicionesdetdoc.idprov)=" & RstFrm("id") & ") AND ((con_provicionesdetdoc.idmod)=" & Rst("id") & "))", xCon
             
            If Rst2.RecordCount <> 0 Then
                Rst2.MoveFirst
                For B = 1 To Rst2.RecordCount
                    If Rst("id") = 1 Then   'compras
                        xCon.Execute "DELETE * FROM com_compras WHERE id = " & Rst2("iddoc") & ""
                    End If
                    If Rst("id") = 2 Then   'Ventas
                        xCon.Execute "DELETE * FROM vta_ventas WHERE id = " & Rst2("iddoc") & ""
                    End If
                    
                    Rst2.MoveNext
                    If Rst2.EOF = True Then Exit For
                Next B
            End If
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
        
        'ELIMINAMOS EL ASIENTO REGISTRADO EN EL DIARIO
        xCon.Execute "DELETE * FROM con_diario WHERE ((idmes = " & mMesActivo & ") and (idlib = 3) AND (idmov = " & RstFrm("id") & ")) ;"
        xCon.Execute "DELETE * FROM con_provicionesdetdoc WHERE idprov = " & RstFrm("id") & ""
        xCon.Execute "DELETE * FROM con_provicionesdet WHERE id = " & RstFrm("id") & ""
        xCon.Execute "DELETE * FROM con_proviciones WHERE id = " & RstFrm("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstFrm("id") & " AND idform = " & IdMenuActivo

        
        xCon.CommitTrans
        RstFrm.Requery
        Dg1.Refresh
        MsgBox "El asiento fue eliminado con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TabOne1.CurrTab = 0
        If RstFrm.RecordCount = 0 Then
            Rpta = MsgBox("No se ha registrado ningún asiento, ¿Desea agregar una ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                Nuevo
            End If
        End If
    End If
    Exit Sub
LaCague:
    xCon.RollbackTrans
    SHOW_ERROR Me.Name, "Eliminar"
End Sub

Private Sub Form_Load()
    Agregando = False
    SeEjecuto = False
    QueHace = 3
    TabOne1.CurrTab = 0
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    CaracteresNumericos = "0123456789." & Chr(8)
    
    Fg1.ColWidth(5) = 0
    Fg1.ColWidth(6) = 0
    Fg1.ColWidth(7) = 0
    Fg1.ColWidth(8) = 0
    
    Dg1.Columns("totdeb1").NumberFormat = FORMAT_MONTO:
    Dg1.Columns("tothab1").NumberFormat = FORMAT_MONTO:
    Dg1.Columns("fchdoc").NumberFormat = FORMAT_DATE:

    Fg1.SelectionMode = flexSelectionByRow
    fg2.SelectionMode = flexSelectionByRow
    
    fg2.ColWidth(8) = 0
    fg2.ColWidth(9) = 0
    fg2.ColWidth(10) = 0
    
    
    fg2.ColWidth(13) = 0
    fg2.ColWidth(14) = 0
    fg2.ColWidth(15) = 0
    fg2.ColWidth(16) = 0
    fg2.ColWidth(17) = 0
    fg2.ColWidth(18) = 0
    fg2.ColWidth(19) = 0
    fg2.ColWidth(20) = 0
    
    Fg1.ColComboList(1) = "|..."
    Fg1.ColComboList(3) = "|..."
    Fg1.ColComboList(4) = "|..."
    
    fg2.ColComboList(2) = "|..."
    fg2.ColComboList(4) = "|..."
    fg2.ColComboList(5) = "|..."
    fg2.ColComboList(6) = "|..."

'    Fg1.Editable = flexEDKbd
End Sub

Private Sub Menu1_1_Click()
    CmdAdd_Click
End Sub

Private Sub Menu1_3_Click()
    CmdDel_Click
End Sub

Private Sub OptDebe_Click()
    If OptDebe.Value = True Then
        Fg1.TextMatrix(Fg1.Row, 7) = 1   ' especifica que la cuenta va al debe
    End If
End Sub

Private Sub OptHaber_Click()
    If OptHaber.Value = True Then
        Fg1.TextMatrix(Fg1.Row, 7) = 2   ' especifica que la cuenta va al haber
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then
            If RstFrm.RecordCount = 0 Then
                Cancel = True
                Exit Sub
            End If
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
            Cancelar2
            RstFrm.Requery
            Dg1.Refresh
            
            If RstFrm.RecordCount <> 0 Then
                RstFrm.MoveFirst
                RstFrm.Find "id=" & mIdRegistro
                If RstFrm.EOF = True Then RstFrm.MoveFirst
            End If
            
        End If
    End If
    If Button.Index = 6 Then Cancelar2
    If Button.Index = 8 Then Filtrar
    If Button.Index = 9 Then
        If xCon.State = 0 Then Exit Sub
        If RstFrm.State = 0 Then Exit Sub
        TDB_FiltroLimpiar Dg1
        RstFrm.Filter = ""
    End If
    If Button.Index = 10 Then CambiarMes
    If Button.Index = 11 Then Buscar
    
    If Button.Index = 13 Then
        If TabOne1.CurrTab = 0 Then IMPRIMIR 1, 0, True
        If TabOne1.CurrTab = 1 Then IMPRIMIR 2, 0, True
    End If
    
    If Button.Index = 15 Then
        If TabOne1.CurrTab = 0 Then
            MsgBox "Para exportar el registro, primero muestre el detalle", vbExclamation, xTitulo
            Exit Sub
        End If
        pExportarMSExcel
    End If
    
    If Button.Index = 17 Then
        Set RstFrm = Nothing
        Unload Me
    End If
    
End Sub

Sub Cancelar2()
    QueHace = 3
    Fg1.SelectionMode = flexSelectionByRow
    Bloquea
    ActivaTool
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Dg1.SetFocus
End Sub

Function Grabar() As Boolean
    If fValidarDatos() = False Then Exit Function
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modficar") + " la Provición", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo) = vbNo Then Exit Function
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstDetDoc As New ADODB.Recordset
'''    Dim RstDia As New ADODB.Recordset
    Dim xNumAsiento As String
    Dim A, B As Integer
    Dim xId2 As Double
    
    On Error GoTo LaCague
    Me.MousePointer = vbHourglass
    
    xCon.BeginTrans
    If QueHace = 1 Then
'''        xNumAsiento = NuevoNumAsiento(3, mMesActivo, xCon)
        xId2 = HallaCodigoTabla("con_proviciones", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM con_proviciones", xCon
        RstCab.AddNew
        RstCab("id") = xId2
    Else
        xId2 = RstFrm("id")
        RST_Busq RstCab, "SELECT * FROM con_proviciones WHERE id = " & xId & "", xCon
        
        'ELIMINAMOS EL DETALLE DE LA PROVICION
        xCon.Execute "DELETE * FROM con_provicionesdetdoc WHERE idprov = " & xId & ""
        xCon.Execute "DELETE * FROM con_provicionesdet WHERE id = " & xId & ""
        
'''        xNumAsiento = DevuelveNumAsiento(3, RstFrm("id"), mMesActivo, xCon)
'''        If xNumAsiento = "" Then xNumAsiento = NuevoNumAsiento(3, mMesActivo, xCon)
'''        'ELIMINAMOS EL ASIENTO REGISTRADO EN EL DIARIO
'''        xCon.Execute "DELETE * FROM con_diario WHERE ((idmes = " & mMesActivo & ") and (idlib = 3) AND (idmov = " & xId & ")) ;"
    End If
    
    mIdRegistro = xId2
    
    RST_Busq RstDet, "SELECT TOP 1 * FROM con_provicionesdet", xCon
    RST_Busq RstDetDoc, "SELECT TOP 1 * FROM con_provicionesdetdoc", xCon
'''    RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
    
    RstCab("ano") = AnoTra
    RstCab("idmes") = mMesActivo
'''    RstCab("numreg") = Format(mMesActivo, "00") + xNumAsiento
    If mMesActivo <> 0 And mMesActivo <> 13 Then
        RstCab("fchreg") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
    End If
    RstCab("idlib") = 3 '--proviciones diversas (libro diario)
    RstCab("idsublib") = NulosN(lbl_cb_cod(1).Caption)
    RstCab("idmon") = NulosN(lbl_cb_cod(0).Caption)
    RstCab("fchdoc") = CDate(TxtFchEmi.Valor)
    RstCab("tipdoc") = NulosN(TxtTipDoc.Text)
    RstCab("numser") = NulosC(TxtSerDoc.Text)
    RstCab("numdoc") = NulosC(TxtNumDoc.Text)
    
    If NulosN(lbl_cb_cod(0).Caption) = 1 Then '--soles
        RstCab("imp") = NulosN(TxtTotDeb.Text)
    Else
        RstCab("imp") = NulosN(TxtTotDebDol.Text)
    End If
    
    RstCab("glosa") = NulosC(TxtGlosa.Text)
    
    RstCab("tc") = NulosN(TxtTC.Text)
    
    '--Especifica si es ajuste x dif cambio
    If ChkAjusteDifCambio.Value = 1 Then
        RstCab("ajuste") = NulosN(lbl_cb_cod(0).Caption)
    Else
        RstCab("ajuste") = 0
    End If
    
    
    RstCab.Update
            
    For A = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("id") = xId2
        RstDet("idcuen") = NulosN(Fg1.TextMatrix(A, 5))
        
        If NulosN(lbl_cb_cod(0).Caption) = 1 Then '--soles
        
            If NulosN(Fg1.TextMatrix(A, 3)) <> 0 Then
                RstDet("tipo") = 0 '--debe
                RstDet("imp") = NulosN(Fg1.TextMatrix(A, 3))
            End If
            
            If NulosN(Fg1.TextMatrix(A, 4)) <> 0 Then
                RstDet("tipo") = -1 '--haber
                RstDet("imp") = NulosN(Fg1.TextMatrix(A, 4))
            End If
        Else '--dolares
            If NulosN(Fg1.TextMatrix(A, 9)) <> 0 Then
                RstDet("tipo") = 0 '--debe
                RstDet("imp") = NulosN(Fg1.TextMatrix(A, 9))
            End If
            
            If NulosN(Fg1.TextMatrix(A, 10)) <> 0 Then
                RstDet("tipo") = -1 '--haber
                RstDet("imp") = NulosN(Fg1.TextMatrix(A, 10))
            End If
        End If
        
        
        
        RstDet.Update
        'si la cuenta tiene detalle
        RstTmp.Filter = adFilterNone
        If RstTmp.RecordCount <> 0 Then
            RstTmp.MoveFirst
            
            'seleccionamos los documentos antiguos para buscar el id en la tabla de documentos que corresponde
            RstTmp.Filter = "idprovi = " & xId & " AND idcuent = " & NulosN(Fg1.TextMatrix(A, 5)) & " AND nuevo = 0"
            If RstTmp.RecordCount <> 0 Then
                RstTmp.MoveFirst
                For B = 1 To RstTmp.RecordCount
                    
                    RstDetDoc.AddNew
                    RstDetDoc("idprov") = xId2
                    RstDetDoc("idcuen") = NulosN(Fg1.TextMatrix(A, 5))
                    RstDetDoc("idmod") = NulosN(Fg1.TextMatrix(A, 9))
                    RstDetDoc("iddoc") = RstTmp("iddoc")
                    
                    RstDetDoc.Update
                    RstTmp.MoveNext
                    If RstTmp.EOF = True Then Exit For
                Next B
            End If
        End If
    Next A
    
    'escribimos los documentos de la provicion en su respectivo modulo
    Dim Rst As New ADODB.Recordset
    Dim xIdDoc As Integer
    Dim RstGraDocMod As New ADODB.Recordset 'recorset para grabar los documentos es su respectivo modulo
    
    If QueHace = 1 Then  'si es nuevo asiento
        For A = 1 To Fg1.Rows - 1
            If NulosN(Fg1.TextMatrix(A, 8)) = -1 Then
                
                If RstTmp.RecordCount <> 0 Then
                
                    RstTmp.Filter = adFilterNone
                    RstTmp.MoveFirst
                    RstTmp.Filter = "idprovi = " & xId & " AND  idcuent  = " & Fg1.TextMatrix(A, 5) & ""
                    
                    If NulosN(Fg1.TextMatrix(A, 6)) = 1 Then ' si es compras
                        RST_Busq RstGraDocMod, "SELECT TOP 1 * FROM com_compras", xCon
                        
                        RstTmp.MoveFirst
                        xIdDoc = HallaCodigoTabla("com_compras", xCon, "id")
                        
                        For B = 1 To RstTmp.RecordCount
                            RstGraDocMod.AddNew
                            RstGraDocMod("id") = xIdDoc
                            RstGraDocMod("idlib") = 1
                            RstGraDocMod("idtipo") = 5
                            RstGraDocMod("tipdoc") = RstTmp("tipdoc")
                            RstGraDocMod("idpro") = RstTmp("idpro")
                            RstGraDocMod("numser") = Mid(RstTmp("numdoc"), 1, 4)
                            RstGraDocMod("numdoc") = Mid(RstTmp("numdoc"), 6, 10)
                            RstGraDocMod("fchreg") = CDate("01/01/08")
                            RstGraDocMod("fchdoc") = RstTmp("fchdoc")
                            RstGraDocMod("fchven") = RstTmp("fchven")
                            RstGraDocMod("idconpag") = RstTmp("idcon")
                            RstGraDocMod("idmon") = RstTmp("idmon")
                            RstGraDocMod("impbru") = RstTmp("impbru")
                            RstGraDocMod("impina") = 0
                            RstGraDocMod("impisc") = RstTmp("impisc")
                            RstGraDocMod("impigv") = RstTmp("impigv")
                            RstGraDocMod("imptot") = RstTmp("imptot")
                            RstGraDocMod("impsal") = RstTmp("impsal")
                            RstGraDocMod("numreg") = Format(mMesActivo, "00") + xNumAsiento
    
                            RstGraDocMod.Update
                            
                            RstDetDoc.AddNew
                            RstDetDoc("idprov") = xId2
                            RstDetDoc("idcue") = Fg1.TextMatrix(A, 5)
                            RstDetDoc("idmod") = 1
                            RstDetDoc("iddoc") = xIdDoc
                            
                            RstDetDoc.Update
                            
                            RstTmp.MoveNext
                            If RstTmp.EOF = True Then Exit For
                            xIdDoc = xIdDoc + 1
                        Next B
                    End If
                    
                    If NulosN(Fg1.TextMatrix(A, 6)) = 2 Then ' si es ventas
                        RST_Busq RstGraDocMod, "SELECT TOP 1 * FROM vta_ventas", xCon
                        
                        RstTmp.MoveFirst
                        xIdDoc = HallaCodigoTabla("vta_ventas", xCon, "id")
                                      
                        For B = 1 To RstTmp.RecordCount
                            RstGraDocMod.AddNew
                            RstGraDocMod("id") = xIdDoc
                            RstGraDocMod("idlib") = 2
                            RstGraDocMod("idtipo") = 2
                            RstGraDocMod("idcli") = RstTmp("idpro")
                            RstGraDocMod("idpunvencli") = 0
                            RstGraDocMod("tipdoc") = RstTmp("tipdoc")
                            RstGraDocMod("numser") = Mid(RstTmp("numdoc"), 1, 4)
                            RstGraDocMod("numdoc") = Mid(RstTmp("numdoc"), 6, 10)
                            RstGraDocMod("fchreg") = CDate("01/01/07")
                            RstGraDocMod("fchdoc") = RstTmp("fchdoc")
                            RstGraDocMod("fchven") = RstTmp("fchven")
                            RstGraDocMod("idconpag") = RstTmp("idcon")
                            RstGraDocMod("idmon") = RstTmp("idmon")
                            RstGraDocMod("impbru") = RstTmp("impbru")
                            RstGraDocMod("impinaf") = 0
                            RstGraDocMod("impisc") = RstTmp("impisc")
                            RstGraDocMod("impigv") = RstTmp("impigv")
                            RstGraDocMod("imptotdoc") = RstTmp("imptot")
                            RstGraDocMod("impsal") = RstTmp("imptot")
                            
                            RstGraDocMod("numreg") = Format(mMesActivo, "00") + xNumAsiento
                            RstGraDocMod("tipgen") = 1
                            RstGraDocMod.Update
                            
                            RstDetDoc.AddNew
                            RstDetDoc("idprov") = xId2
                            RstDetDoc("idcuen") = Fg1.TextMatrix(A, 5)
                            RstDetDoc("idmod") = 2
                            RstDetDoc("iddoc") = xIdDoc
                            
                            RstDetDoc.Update
                            
                            RstTmp.MoveNext
                            If RstTmp.EOF = True Then Exit For
                            xIdDoc = xIdDoc + 1
                        
                        Next B
                    End If
                End If
                If NulosN(Fg1.TextMatrix(A, 6)) = 4 Then ' si es ventas
                End If
                
                If NulosN(Fg1.TextMatrix(A, 6)) = 8 Then ' si es ventas
                End If
            End If
        Next A
    End If
    
    
'    xCampos(0, 0) = "fchdoc":     xCampos(0, 1) = "C":      xCampos(0, 2) = "10"  'fecha del documento
'    xCampos(1, 0) = "numdoc":     xCampos(1, 1) = "C":      xCampos(1, 2) = "15" 'numero del documento
'    xCampos(2, 0) = "nommon":     xCampos(2, 1) = "C":      xCampos(2, 2) = "20" 'descrpcion de la moneda
'    xCampos(3, 0) = "provee":     xCampos(3, 1) = "C":      xCampos(3, 2) = "100"  'nombre del proveedor
'    xCampos(4, 0) = "nomcon":     xCampos(4, 1) = "C":      xCampos(4, 2) = "20" 'descripcion de la condicion de venta
'    xCampos(5, 0) = "fchven":     xCampos(5, 1) = "F":      xCampos(5, 2) = "10" 'fecha de vencimiento
'    xCampos(6, 0) = "impbru":     xCampos(6, 1) = "D":      xCampos(6, 2) = "20" 'importe bruto
'    xCampos(7, 0) = "impigv":     xCampos(7, 1) = "D":      xCampos(7, 2) = "20" 'importe del igv
'    xCampos(8, 0) = "impisc":     xCampos(8, 1) = "D":      xCampos(8, 2) = "20" 'importe del isc
'    xCampos(9, 0) = "imptot":     xCampos(9, 1) = "D":      xCampos(9, 2) = "20" ' importe total del documento
'    xCampos(10, 0) = "impsal":    xCampos(10, 1) = "D":     xCampos(10, 2) = "20" 'importe del saldo del documento
'    xCampos(11, 0) = "iddoc":     xCampos(11, 1) = "N":     xCampos(11, 2) = "2" 'id del documento que se esta cargando
'    xCampos(12, 0) = "idmon":     xCampos(12, 1) = "N":     xCampos(12, 2) = "2" 'id de la moneda del documento
'    xCampos(13, 0) = "idcon":     xCampos(13, 1) = "N":     xCampos(13, 2) = "2" 'id de la condicion del documento
'    xCampos(14, 0) = "idpro":     xCampos(14, 1) = "N":     xCampos(14, 2) = "2" 'id del proveedor
'    xCampos(15, 0) = "idprovi":   xCampos(15, 1) = "N":     xCampos(15, 2) = "2" 'id de la condicion del documento
'    xCampos(16, 0) = "idcuent":   xCampos(16, 1) = "N":     xCampos(16, 2) = "2" 'id del proveedor
'    xCampos(17, 0) = "nuevo":     xCampos(17, 1) = "N":     xCampos(17, 2) = "2" 'especifica si es un documento nuevo
'    xCampos(18, 0) = "edita":     xCampos(18, 1) = "N":     xCampos(18, 2) = "2" 'especifica si el documento se  puede modificar
'    xCampos(19, 0) = "tipdoc":    xCampos(19, 1) = "N":     xCampos(19, 2) = "2" 'especifica si el documento se  puede modificar
    
    If QueHace = 2 Then  'si se modifica un asiento
    End If
    
'    For A = 1 To Fg1.Rows - 1
'        If NulosN(Fg1.TextMatrix(A, 6)) = 1 Then ' si es compras
'            'cargamos los documentos de la cuenta actual
'            RstTmp.Filter = adFilterNone
'            RstTmp.MoveFirst
'            RstTmp.Filter = "idprovi = " & xId & " AND  idcuent  = " & Fg1.TextMatrix(A, 5) & ""
'
'            If RstTmp.RecordCount <> 0 Then
'                'eliminamos lo documentos de com_compras que no esten en el rsttmp
'                RST_Busq Rst, "SELECT com_compras.* From com_compras WHERE (((Mid([numreg],1,2))='" & Format(mMesActivo, "00") & "'))", xCon
'                If Rst.RecordCount <> 0 Then
'                    Rst.MoveFirst
'                    For B = 1 To Rst.RecordCount
'                        RstTmp.MoveFirst
'                        RstTmp.Find "iddoc = " & Rst("id") & ""
'                        If RstTmp.EOF = True Then
'                            'no existe el documento en el rst RSTTMP, lo borramos de la tabla com_compras
'                            xCon.Execute "DELETE * FROM com_compras WHERE id = " & Rst("id") & ""
'                        End If
'
'                        Rst.MoveNext
'                        If Rst.EOF = True Then Exit For
'                    Next B
'                    Rst.MoveFirst
'                End If
'            End If
'
'            'cargamos los documentos de la cuenta actual, solo se cargan los docuentos que se pueden editar
'            RstTmp.Filter = adFilterNone
'            RstTmp.MoveFirst
'            RstTmp.Filter = "idprovi = " & xId & " AND  idcuent  = " & Fg1.TextMatrix(A, 5) & " And edita = 1"
'
'            If RstTmp.RecordCount <> 0 Then
'                RstTmp.MoveFirst
'                For B = 1 To RstTmp.RecordCount
'                    'borramos los documentos de la tablas con_compras que esten en el RSTTMP
'                    xCon.Execute "DELETE * FROM com_compras WHERE id = " & RstTmp("idddoc") & ""
'                    RstTmp.MoveNext
'                    If RstTmp.EOF = True Then Exit For
'                Next B
'            End If
'        End If
'    Next A
'
''    xCampos(0, 0) = "fchdoc":     xCampos(0, 1) = "C":      xCampos(0, 2) = "10"  'fecha del documento
''    xCampos(1, 0) = "numdoc":     xCampos(1, 1) = "C":      xCampos(1, 2) = "15" 'numero del documento
''    xCampos(2, 0) = "nommon":     xCampos(2, 1) = "C":      xCampos(2, 2) = "20" 'descrpcion de la moneda
''    xCampos(3, 0) = "provee":     xCampos(3, 1) = "C":      xCampos(3, 2) = "100"  'nombre del proveedor
''    xCampos(4, 0) = "nomcon":     xCampos(4, 1) = "C":      xCampos(4, 2) = "20" 'descripcion de la condicion de venta
''    xCampos(5, 0) = "fchven":     xCampos(5, 1) = "F":      xCampos(5, 2) = "10" 'fecha de vencimiento
''    xCampos(6, 0) = "impbru":     xCampos(6, 1) = "D":      xCampos(6, 2) = "20" 'importe bruto
''    xCampos(7, 0) = "impigv":     xCampos(7, 1) = "D":      xCampos(7, 2) = "20" 'importe del igv
''    xCampos(8, 0) = "impisc":     xCampos(8, 1) = "D":      xCampos(8, 2) = "20" 'importe del isc
''    xCampos(9, 0) = "imptot":     xCampos(9, 1) = "D":      xCampos(9, 2) = "20" ' importe total del documento
''    xCampos(10, 0) = "impsal":    xCampos(10, 1) = "D":     xCampos(10, 2) = "20" 'importe del saldo del documento
''    xCampos(11, 0) = "iddoc":     xCampos(11, 1) = "N":     xCampos(11, 2) = "2" 'id del documento que se esta cargando
''    xCampos(12, 0) = "idmon":     xCampos(12, 1) = "N":     xCampos(12, 2) = "2" 'id de la moneda del documento
''    xCampos(13, 0) = "idcon":     xCampos(13, 1) = "N":     xCampos(13, 2) = "2" 'id de la condicion del documento
''    xCampos(14, 0) = "idpro":     xCampos(14, 1) = "N":     xCampos(14, 2) = "2" 'id del proveedor
''    xCampos(15, 0) = "idprovi":   xCampos(15, 1) = "N":     xCampos(15, 2) = "2" 'id de la condicion del documento
''    xCampos(16, 0) = "idcuent":   xCampos(16, 1) = "N":     xCampos(16, 2) = "2" 'id del proveedor
''    xCampos(17, 0) = "nuevo":     xCampos(17, 1) = "N":     xCampos(17, 2) = "2" 'especifica si es un documento nuevo
''    Set RstTmp = xFun.CrearRstTMP(xCampos)
    
    'grabamos el diario
'''''''    For A = 1 To Fg1.Rows - 1
'''''''        RstDia.AddNew
'''''''        RstDia("año") = AnoTra
'''''''        RstDia("idmes") = mMesActivo
'''''''        RstDia("idlib") = 3 'NulosN(TxtIdLibro.Text)
'''''''        RstDia("idmov") = xId2
'''''''        RstDia("idcue") = NulosN(Fg1.TextMatrix(A, 5))
'''''''        RstDia("numasi") = xNumAsiento
'''''''        RstDia("tc") = NulosN(LblTipoCambio.Caption)
'''''''
'''''''        RstDia("impdebsol") = NulosN(Fg1.TextMatrix(A, 3))
'''''''        RstDia("imphabsol") = NulosN(Fg1.TextMatrix(A, 4))
'''''''
'''''''        RstDia("impdebdol") = NulosN(Fg1.TextMatrix(A, 9))
'''''''        RstDia("imphabdol") = NulosN(Fg1.TextMatrix(A, 10))
'''''''
'''''''        If mMesActivo = 13 Then
'''''''            RstDia("fchasi") = CDate("31/12/" + AnoTra)
'''''''        Else
'''''''            If mMesActivo = 0 Then
'''''''                RstDia("fchasi") = CDate("31/12/" + Str(AnoTra - 1))
'''''''            Else
'''''''                RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''''''            End If
'''''''        End If
'''''''        RstDia("fchdoc") = CDate(TxtFchEmi.Valor)
'''''''        RstDia("prodiv") = -1
'''''''        RstDia.Update
'''''''    Next A
    '----------------------------------------------------------------------------------
    '---generar asiento
    xNumAsiento = GenerarAsiento(xCon, 3, xId2, AnoTra, mMesActivo, 1)
    If xNumAsiento = "" Then GoTo LaCague
    '----------------------------------------------------------------------------------
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId2
    
    xCon.CommitTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
'''    Set RstDia = Nothing
    MsgBox "La Provición se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito" + vbCr + "Num.Reg. " & xNumAsiento, vbInformation, xTitulo
'''    MsgBox "La Provición se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito" + vbCr + "Num.Reg. " & Format(mMesActivo, "00") & xNumAsiento, vbInformation, xTitulo

    Grabar = True
    Me.MousePointer = vbDefault
    Exit Function
    
LaCague:
    'Resume
    Me.MousePointer = vbDefault
    Grabar = False
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
'''    Set RstDia = Nothing
    MsgBox "No se pudo guardar la provicion por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Function

Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)
    If txt_cb(Index).Text = "" Then Exit Sub
    If QueHace = 3 Then Exit Sub
    Dim RST_TMP As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
        Case 0 '--MONEDA
            nSQL = "SELECT mae_moneda.id, mae_moneda.descripcion AS nombre, mae_moneda.id AS cod FROM mae_moneda WHERE (((mae_moneda.id)=" & NulosN(Trim(txt_cb(Index).Text)) & "))"
        Case 1 '--SUB LIBRO
            If NulosN(TxtIdLibro.Text) = 0 Then
                MsgBox "Falta especificar el libro", vbExclamation, xTitulo
                Exit Sub
            End If
            nSQL = "SELECT mae_librossub.id, mae_librossub.descripcion AS nombre, mae_librossub.id AS cod " _
                + vbCr + " FROM mae_librossub " _
                + vbCr + " WHERE (((mae_librossub.idlib) = " & NulosN(TxtIdLibro.Text) & ")) " _
                + vbCr + " AND (((mae_librossub.id)=" & NulosN(Trim(txt_cb(Index).Text)) & ")) ;"
                
    End Select
    If xCon.State = 0 Then Exit Sub
    RST_Busq RST_TMP, nSQL, xCon
    
    If RST_TMP.State = 0 Then GoTo SALIR
    If RST_TMP.RecordCount > 0 Then
        txt_cb(Index).Text = RST_TMP.Fields(0) & ""  '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RST_TMP.Fields(1) & "" '--NOMBRE
        lbl_cb_cod(Index).Caption = RST_TMP.Fields(2) & "" '--CODIGO
        
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cb_cod(Index).Caption = ""
        
    End If
    
SALIR:
    Set RST_TMP = Nothing
    Exit Sub
error:
    Set RST_TMP = Nothing
    SHOW_ERROR Me.Name, "txt_cb_KeyDown (" + CStr(Index) + ")"
End Sub

Private Sub TxtFchEmi_Validate(Cancel As Boolean)
    
    If IsDate(TxtFchEmi.Valor) = True Then
        If mMesActivo = 0 Then
            'obtenemos el tipo de cambio del 31/12/xxxx para hacer la conversion a dolares, esto se hace solo para asientos de apertura
            Dim xMes1  As String
            xMes1 = "01/01/" + AnoTra
            LblTipoCambio.Caption = HallaTipoCambio(CDate(xMes1) - 1, 2, Venta, xCon)
            If NulosN(LblTipoCambio.Caption) = 0 Then
                MsgBox "No se ha especificado el tipo de cambio para el dia " & CDate(xMes1) - 1 & vbCr & "Ingrese el tipo de cambio para seguir con esta operación" + vbCr + "Ir a Menu Contabilidad/Tipo de Cambio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                TxtFchEmi.Valor = ""
                LblTipoCambio.Caption = ""
                TxtFchEmi.SetFocus
                Exit Sub
            End If
        Else
            LblTipoCambio.Caption = HallaTipoCambio(TxtFchEmi.Valor, 2, Venta, xCon)
        End If
    Else
        LblTipoCambio.Caption = ""
    End If
    
    TxtTC.Text = Format(NulosN(LblTipoCambio.Caption), "0.000")
    pRecalcularImporte
End Sub

Private Sub TxtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Fg1.Rows > 1 Then
             Fg1.Row = 1
             Fg1.Col = 1
             Fg1.SetFocus
        Else
            CmdAdd.SetFocus
        End If
    End If
End Sub

Private Sub TxtIdLibro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdLibro_Validate(Cancel As Boolean)
    If NulosC(TxtIdLibro.Text) = "" Then Exit Sub
    LblLibro.Caption = Busca_Codigo(NulosN(TxtIdLibro.Text), "id", "descripcion", "mae_libros", "N", xCon)
    If LblLibro.Caption = "" Then
        TxtIdLibro.Text = ""
    End If
End Sub

Private Sub TxtIdMod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtIdMod_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMod_Click
    End If
End Sub

Private Sub TxtIdMod_Validate(Cancel As Boolean)
    If NulosN(TxtIdMod.Text) <> 0 Then
        LblDescModulo.Caption = Busca_Codigo(NulosN(TxtIdMod.Text), "id", "descripcion", "tes_modulos", "N", xCon)
        If NulosC(LblDescModulo.Caption) = "" Then
            TxtIdMod.Text = ""
        Else
            Fg1.TextMatrix(Fg1.Row, 6) = TxtIdMod.Text
        End If
    End If
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumDoc_Validate(Cancel As Boolean)
    If TxtNumDoc.Text = "" Then Exit Sub
    TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
End Sub

Private Sub TxtSerDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtSerDoc_Validate(Cancel As Boolean)
    If TxtSerDoc.Text = "" Then Exit Sub
    TxtSerDoc.Text = Format(TxtSerDoc.Text, "0000")
End Sub

Private Sub TxtTC_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTC_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    TxtTC.Text = Format(TxtTC.Text, "0.000")
    pRecalcularImporte
End Sub

Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipDoc_Validate(Cancel As Boolean)
    If NulosC(TxtTipDoc.Text) = "" Then Exit Sub
    
    LblDocumento.Caption = Busca_Codigo(Val(TxtTipDoc.Text), "id", "descripcion", "mae_documento", "N", xCon)
    If LblDocumento.Caption = "" Then
        TxtTipDoc.Text = ""
        LblDocumento.Caption = ""
    End If
End Sub

Sub MuestraSegundoTab()
'    On Error GoTo error
    Fg1.Rows = 1
    Blanquea
    If RstFrm.BOF = True Or RstFrm.EOF = True Then Exit Sub
    If RstFrm.RecordCount = 0 Then Exit Sub
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer
    QueHace = -1

    PreparaRST_Tmp
    
    lblReg.Caption = "Nº Reg. " & NulosC(RstFrm("registro"))
    
    TxtIdLibro.Text = NulosN(RstFrm("idlib"))
    LblLibro.Caption = NulosC(RstFrm("desclib"))
    'LblDocumento.Caption = NulosC(RstFrm("descdoc"))
    TxtSerDoc.Text = NulosC(RstFrm("numser"))
    TxtNumDoc.Text = NulosC(RstFrm("numdoc"))
    TxtFchEmi.Valor = NulosC(RstFrm("fchdoc"))
    TxtFchEmi_Validate False
    
    TxtGlosa.Text = NulosC(RstFrm("glosa"))
       
    If NulosN(RstFrm.Fields("idsublib")) <> 0 Then
        txt_cb(1).Text = RstFrm.Fields("idsublib")
        txt_cb_Validate 1, False
    End If
    
    If NulosN(RstFrm.Fields("idmon")) <> 0 Then
        txt_cb(0).Text = RstFrm.Fields("idmon")
        txt_cb_Validate 0, False
    End If
    If NulosN(RstFrm("tipdoc")) <> 0 Then
        TxtTipDoc.Text = NulosN(RstFrm("tipdoc"))
        TxtTipDoc_Validate False
    End If
    
    TxtTC.Text = NulosN(RstFrm("tc"))
    
    '--Especifica si es ajuste x dif cambio
    If NulosN(RstFrm("ajuste")) <> 0 Then
        ChkAjusteDifCambio.Value = 1
    Else
        ChkAjusteDifCambio.Value = 0
    End If
    
    QueHace = 3
   
    RST_Busq RstDet, "SELECT con_provicionesdet.id, con_provicionesdet.idcuen, con_planctas.cuenta, con_planctas.descripcion, " _
        & " IIf([con_provicionesdet].[tipo]=0,[con_provicionesdet]![imp],0) AS debe, IIf([con_provicionesdet].[tipo]=-1,[con_provicionesdet]![imp],0) AS haber, " _
        & " (SELECT Count([idcuen]) AS numdoc From con_provicionesdetdoc WHERE (((con_provicionesdetdoc.idprov)=con_provicionesdet.id) AND  " _
        & " ((con_provicionesdetdoc.idcuen)=con_provicionesdet.idcuen))) AS cantdoc, con_planctas.idmodulo FROM con_provicionesdet LEFT JOIN con_planctas " _
        & " ON con_provicionesdet.idcuen = con_planctas.id Where (((con_provicionesdet.id) = " & NulosN(RstFrm("id")) & ")) ORDER BY con_planctas.cuenta", xCon
    
    Agregando = True
    
    If RstDet.RecordCount <> 0 Then
        RstDet.MoveFirst
        For A = 1 To RstDet.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.Row = Fg1.Rows - 1
            Fg1.TextMatrix(A, 1) = NulosC(RstDet("cuenta"))
            Fg1.TextMatrix(A, 2) = NulosC(RstDet("descripcion"))
            
            Fg1.TextMatrix(A, 3) = 0
            Fg1.TextMatrix(A, 4) = 0
            Fg1.TextMatrix(A, 9) = 0
            Fg1.TextMatrix(A, 10) = 0
            
            If NulosN(RstFrm.Fields("idmon")) = 1 Then
                Fg1.TextMatrix(A, 3) = NulosN(RstDet("debe"))
                Fg1.TextMatrix(A, 4) = NulosN(RstDet("haber"))
                If ChkAjusteDifCambio.Value = 0 Then
                    If NulosN(TxtTC.Text) <> 0 Then '--en dolares
                        Fg1.TextMatrix(A, 9) = NulosN(Fg1.TextMatrix(A, 3)) / NulosN(TxtTC.Text)
                        Fg1.TextMatrix(A, 10) = NulosN(Fg1.TextMatrix(A, 4)) / NulosN(TxtTC.Text)
                    End If
                End If
            Else
                Fg1.TextMatrix(A, 9) = NulosN(RstDet("debe"))
                Fg1.TextMatrix(A, 10) = NulosN(RstDet("haber"))
                
                If ChkAjusteDifCambio.Value = 0 Then
                    '--en soles
                    Fg1.TextMatrix(A, 3) = NulosN(Fg1.TextMatrix(A, 9)) * NulosN(TxtTC.Text)
                    Fg1.TextMatrix(A, 4) = NulosN(Fg1.TextMatrix(A, 10)) * NulosN(TxtTC.Text)
                End If
                
            End If
            
            Fg1.TextMatrix(A, 3) = Format(Fg1.TextMatrix(A, 3), FORMAT_MONTO)
            Fg1.TextMatrix(A, 4) = Format(Fg1.TextMatrix(A, 4), FORMAT_MONTO)
            Fg1.TextMatrix(A, 9) = Format(Fg1.TextMatrix(A, 9), FORMAT_MONTO)
            Fg1.TextMatrix(A, 10) = Format(Fg1.TextMatrix(A, 10), FORMAT_MONTO)
            
            Agregando = True
            
            Fg1.TextMatrix(A, 5) = NulosN(RstDet("idcuen"))
            Fg1.TextMatrix(A, 6) = NulosN(RstDet("idmodulo"))
            Fg1.TextMatrix(A, 7) = NulosN(RstDet("cantdoc"))
            If RstDet("idmodulo") = 1 Then   'Cargamos las compras
                AgregarDocCompras RstFrm("id"), NulosN(RstDet("idcuen"))
            End If
            
            If RstDet("idmodulo") = 2 Then   'Cargamos las ventas
                AgregarDocVentas RstFrm("id"), RstDet("idcuen")
            End If
            
            RstDet.MoveNext
            If RstDet.EOF = True Then Exit For
        Next A
    End If
    
    
    TxtTotDeb.Text = Format(SumaColumna(Fg1, 3), FORMAT_MONTO)
    TxtTotHab.Text = Format(SumaColumna(Fg1, 4), FORMAT_MONTO)
    TxtTotDebDol.Text = Format(SumaColumna(Fg1, 9), FORMAT_MONTO)
    TxtTotHabDol.Text = Format(SumaColumna(Fg1, 10), FORMAT_MONTO)
    
    Agregando = False
    Exit Sub

error:
    Resume
    Agregando = False
    SHOW_ERROR Me.Name, "MuestraSegundoTab"
End Sub

Function TieneMovimientosDoc(IdDocumento As Integer, IdModulo As Integer) As Boolean
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT tes_cajadestinodet.iddoc, tes_cajadestinodet.idmod, tes_cajadestinodet.acuenta From tes_cajadestinodet " _
        & " WHERE (((tes_cajadestinodet.iddoc)=" & IdDocumento & ") AND ((tes_cajadestinodet.idmod)= " & IdModulo & ") AND ((tes_cajadestinodet.acuenta)<>0))", xCon

    If Rst.RecordCount <> 0 Then
        TieneMovimientosDoc = True
    Else
        TieneMovimientosDoc = False
    End If
    Set Rst = Nothing
End Function

Sub AgregarDocCompras(IdProvicion As Integer, IdCuenta As Integer)
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    
    RST_Busq Rst, "SELECT com_compras.*, mae_moneda.simbolo, mae_prov.nombre, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdocu, mae_condpago.abrev " _
        & " FROM (con_provicionesdetdoc LEFT JOIN ((mae_moneda RIGHT JOIN com_compras ON mae_moneda.id = com_compras.idmon) LEFT JOIN mae_prov " _
        & " ON com_compras.idpro = mae_prov.id) ON con_provicionesdetdoc.iddoc = com_compras.id) LEFT JOIN mae_condpago ON com_compras.idconpag = mae_condpago.id " _
        & " WHERE (((con_provicionesdetdoc.idprov)=" & IdProvicion & ") AND ((con_provicionesdetdoc.idcuen)=" & IdCuenta & "))", xCon
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            RstTmp.AddNew
            RstTmp("fchdoc") = Rst("fchdoc")
            RstTmp("numdoc") = Rst("numdocu")
            RstTmp("nommon") = Rst("simbolo")
            RstTmp("provee") = Rst("nombre")
            RstTmp("nomcon") = Rst("abrev")
            RstTmp("fchven") = Rst("fchven")
            RstTmp("impbru") = Rst("impbru")
            RstTmp("impigv") = Rst("impigv")
            RstTmp("impisc") = Rst("impisc")
            RstTmp("imptot") = Rst("imptot")
            RstTmp("impsal") = Rst("impsal")
            RstTmp("iddoc") = Rst("id")
            RstTmp("idmon") = Rst("idmon")
            RstTmp("idcon") = Rst("idconpag")
            RstTmp("idpro") = Rst("idpro")
            RstTmp("idprovi") = IdProvicion
            RstTmp("idcuent") = IdCuenta
            If TieneMovimientosDoc(Rst("id"), 1) = False Then
                RstTmp("edita") = 0 'no se edita el documento por tener movimiento
            Else
                RstTmp("edita") = 1 'se edita el documento
            End If
            
            Rst.MoveNext
            
            If Rst.EOF = True Then Exit For
        Next A
    End If
    
End Sub

Sub AgregarDocVentas(IdProvicion As Integer, IdCuenta As Integer)
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    
    RST_Busq Rst, "SELECT vta_ventas.*, mae_moneda.simbolo, mae_cliente.nombre, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numdocu, " _
        & " mae_condpago.abrev FROM ((mae_cliente RIGHT JOIN (con_provicionesdetdoc LEFT JOIN vta_ventas ON con_provicionesdetdoc.iddoc = vta_ventas.id) " _
        & " ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN mae_condpago " _
        & " ON vta_ventas.idconpag = mae_condpago.id Where (((con_provicionesdetdoc.idprov) = " & IdProvicion & ") And ((con_provicionesdetdoc.idcuen) = " & IdCuenta & ")) " _
        & " ORDER BY [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc]", xCon

    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            RstTmp.AddNew
            RstTmp("fchdoc") = NulosC(Rst("fchdoc"))
            RstTmp("numdoc") = NulosC(Rst("numdocu"))
            RstTmp("nommon") = NulosC(Rst("simbolo"))
            RstTmp("provee") = NulosC(Rst("nombre"))
            RstTmp("nomcon") = NulosC(Rst("abrev"))
            RstTmp("fchven") = NulosC(Rst("fchven"))
            RstTmp("impbru") = Rst("impbru")
            RstTmp("impigv") = Rst("impigv")
            RstTmp("impisc") = Rst("impisc")
            RstTmp("imptot") = Rst("imptotdoc")
            RstTmp("impsal") = Rst("impsal")
            RstTmp("iddoc") = Rst("id")
            RstTmp("idmon") = Rst("idmon")
            RstTmp("idcon") = Rst("idconpag")
            RstTmp("idpro") = Rst("idcli")
            RstTmp("idprovi") = IdProvicion
            RstTmp("idcuent") = IdCuenta
            Rst.MoveNext
            
            If Rst.EOF = True Then Exit For
        Next A
    End If
    
'    xCampos(0, 0) = "fchdoc":     xCampos(0, 1) = "N":      xCampos(0, 2) = "20"  'fecha del documento
'    xCampos(1, 0) = "numdoc":     xCampos(1, 1) = "N":      xCampos(1, 2) = "200" 'numero del documento
'    xCampos(2, 0) = "nommon":     xCampos(2, 1) = "N":      xCampos(2, 2) = "200" 'descrpcion de la moneda
'    xCampos(3, 0) = "provee":     xCampos(3, 1) = "C":      xCampos(3, 2) = "15"  'nombre del proveedor
'    xCampos(4, 0) = "noncon":     xCampos(4, 1) = "C":      xCampos(4, 2) = "100" 'descripcion de la condicion de venta
'    xCampos(5, 0) = "fchven":     xCampos(5, 1) = "D":      xCampos(5, 2) = "200" 'fecha de vencimiento
'    xCampos(6, 0) = "impbru":     xCampos(6, 1) = "D":      xCampos(6, 2) = "200" 'importe bruto
'    xCampos(7, 0) = "impigv":     xCampos(7, 1) = "D":      xCampos(7, 2) = "200" 'importe del igv
'    xCampos(8, 0) = "impisc":     xCampos(8, 1) = "D":      xCampos(8, 2) = "200" 'importe del isc
'    xCampos(9, 0) = "imptot":     xCampos(9, 1) = "D":      xCampos(9, 2) = "200" ' importe total del documento
'    xCampos(10, 0) = "impsal":    xCampos(10, 1) = "D":     xCampos(10, 2) = "200" 'importe del saldo del documento
'    xCampos(11, 0) = "iddoc":     xCampos(11, 1) = "D":     xCampos(11, 2) = "200" 'id del documento que se esta cargando
'    xCampos(12, 0) = "idmon":     xCampos(12, 1) = "D":     xCampos(12, 2) = "200" 'id de la moneda del documento
'    xCampos(13, 0) = "idcon":     xCampos(13, 1) = "D":     xCampos(13, 2) = "200" 'id de la condicion del documento
'    xCampos(14, 0) = "idpro":     xCampos(14, 1) = "D":     xCampos(14, 2) = "200" 'id del proveedor
    
End Sub


Private Sub CambiarMes()
    mMesActivo = SeleccionaMes(xCon)
    OpcionesPeriodo
    TabOne1.CurrTab = 0
End Sub

Private Function fValidarDatos() As Boolean
    '--VALIDAR QUE LA GRILLA DE ACTIVO Y PASIVO TENGAN VALORES TANTO DE ORDEN Y DESCRIPCION
    
    If NulosN(TxtIdLibro.Text) = 0 Then
        MsgBox "No ha especificado el libro contable", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Function
    End If
    If NulosN(lbl_cb_cod(1).Caption) = 0 Then
        MsgBox "No ha especificado el Sub Libro para el libro " & LblLibro.Caption, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        txt_cb(1).SetFocus
        Exit Function
    End If
    If NulosN(lbl_cb_cod(0).Caption) = 0 Then
        MsgBox "No ha especificado la Moneda", vbInformation, xTitulo
        txt_cb(0).SetFocus
        Exit Function
    End If
    
    If IsDate(TxtFchEmi.Valor) = False Then
        MsgBox "No ha especificado la fecha de movimiento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchEmi.SetFocus
        Exit Function
    End If
    
    If NulosN(TxtTC.Text) = 0 Then
        MsgBox "Falta ingresar el tipo de Cambio", vbExclamation, xTitulo
        TxtTC.SetFocus
        Exit Function
    End If
    
    
    If Trim(TxtGlosa.Text) = "" Then
        MsgBox "No ha especificado la glosa del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtGlosa.SetFocus
        Exit Function
    End If
    
    If Fg1.Rows <= 1 Then
        MsgBox "Ingrese las Cuentas Contables", vbExclamation, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    '--------------------------------
    HallarTotal
    
    If NulosN(txt_cb(0).Text) = 1 Then '--en MN
        If NulosN(TxtTotDeb.Text) <> NulosN(TxtTotHab.Text) Then
            MsgBox "Los totales del Debe y del Haber son diferentes" + vbCr + "Estos tienen que ser iguales", vbExclamation, xTitulo
            Exit Function
        End If
    Else '--en ME
        If NulosN(TxtTotDebDol.Text) <> NulosN(TxtTotHabDol.Text) Then
            MsgBox "Los totales del Debe y del Haber son diferentes" + vbCr + "Estos tienen que ser iguales", vbExclamation, xTitulo
            Exit Function
        End If
    End If
    '--------------------------------
    '--VALIDAR QUE EXISTA VALOR EN DEBE O HABER DE UAN FILA
    '--VALIDAR EL INGRESO DE LOS DATOS
    Dim mRow&
    Dim mCol& '--COLUMNA A POSICIONAR SI FALTAN DATOS
    mCol = -1
    For mRow = 1 To Fg1.Rows - 1
        If Fg1.TextMatrix(mRow, 1) = "" Then
            MsgBox "Ingrese La Cuenta Contable", vbExclamation, xTitulo
            mCol = 1:          Exit For
        ElseIf NulosN(Fg1.TextMatrix(mRow, 3)) = 0 And NulosN(Fg1.TextMatrix(mRow, 4)) = 0 Then
            MsgBox "Ingrese un valor en el Debe o Haber" + vbCr + "Luego Proceda", vbExclamation, xTitulo
            mCol = 3:          Exit For
        End If
    Next mRow
    If mCol <> -1 Then
        Agregando = True:  Fg1.Row = mRow: Fg1.Col = mCol: Agregando = False
        Exit Function
    End If
    '-----
    fValidarDatos = True
End Function
 
Sub Filtrar()
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(6, 4) As String

    xCampos(0, 0) = "Num.Reg.":       xCampos(0, 1) = "registro":      xCampos(0, 2) = "C":          xCampos(0, 3) = "800"
    xCampos(1, 0) = "Libro":           xCampos(1, 1) = "desclib":     xCampos(1, 2) = "C":          xCampos(1, 3) = "1500"
    xCampos(2, 0) = "Descripción":     xCampos(2, 1) = "glosa":       xCampos(2, 2) = "C":          xCampos(2, 3) = "6500"
    xCampos(3, 0) = "M":               xCampos(3, 1) = "mondesc":     xCampos(3, 2) = "C":          xCampos(3, 3) = "450"
    xCampos(4, 0) = "Debe":            xCampos(4, 1) = "totdeb":      xCampos(4, 2) = "N":          xCampos(4, 3) = "1000"
    xCampos(5, 0) = "Haber":           xCampos(5, 1) = "tothab":      xCampos(5, 2) = "N":          xCampos(5, 3) = "1000"

    TabOne1.CurrTab = 0
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg1

End Sub

Sub Buscar()
    On Error GoTo error
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim xCampos2(6, 4) As String
    
    xCampos2(0, 0) = "Num.Reg.":        xCampos2(0, 1) = "registro":      xCampos2(0, 2) = "900":          xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Sub Libro":       xCampos2(1, 1) = "sublibdesc":     xCampos2(1, 2) = "1900":         xCampos2(1, 3) = "C"
    xCampos2(2, 0) = "Descripción":     xCampos2(2, 1) = "glosa":       xCampos2(2, 2) = "3000":         xCampos2(2, 3) = "C"
    xCampos2(3, 0) = "M":               xCampos2(3, 1) = "simbolo":     xCampos2(3, 2) = "450":          xCampos2(3, 3) = "C"
    xCampos2(4, 0) = "Debe":            xCampos2(4, 1) = "totdeb":      xCampos2(4, 2) = "1000":         xCampos2(4, 3) = "N"
    xCampos2(5, 0) = "Haber":           xCampos2(5, 1) = "tothab":      xCampos2(5, 2) = "1000":         xCampos2(5, 3) = "N"
    
    nSQL = "SELECT con_proviciones.*, mae_libros.descripcion AS desclib, mae_documento.abrev AS destipdoc, con_meses.descripcion AS descmes, " _
        + vbCr + " (SELECT Sum([con_provicionesdet]![imp]) AS totdeb  FROM con_provicionesdet WHERE (((con_provicionesdet.id)=con_proviciones.id)  AND ((con_provicionesdet.tipo)=0))) AS totdeb, " _
        + vbCr + " (SELECT Sum([con_provicionesdet]![imp]) AS totdeb FROM con_provicionesdet WHERE  (((con_provicionesdet.id)=con_proviciones.id)   AND ((con_provicionesdet.tipo)=-1))) AS tothab, " _
        + vbCr + " mae_moneda.descripcion AS mondesc, mae_moneda.simbolo, [con_proviciones]![numser]+'-'+[con_proviciones]![numdoc] AS numedoc, mae_librossub.descripcion AS sublibdesc, " _
        + vbCr + " Format([con_proviciones].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & Mid([con_proviciones].[numreg],3) AS registro " _
        + vbCr + " FROM ((((con_proviciones LEFT JOIN mae_libros ON con_proviciones.idlib = mae_libros.id) LEFT JOIN con_meses ON con_proviciones.idmes = con_meses.id) LEFT JOIN mae_moneda ON con_proviciones.idmon = mae_moneda.id) LEFT JOIN mae_documento ON con_proviciones.tipdoc = mae_documento.id) LEFT JOIN mae_librossub ON con_proviciones.idsublib = mae_librossub.id " _
        + vbCr + " Where (((con_proviciones.ano) = " & AnoTra & "  ) And ((con_proviciones.idmes) = " & mMesActivo & "  )) " _
        + vbCr + " ORDER BY con_proviciones.fchreg;"
    
    TabOne1.CurrTab = 0
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos2(), "Buscando Asientos Diversos", "glosa", "glosa", Principio
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    
    RstFrm.MoveFirst
    RstFrm.Find "id = " & xRs("id") & ""
SALIR:
    Set xRs = Nothing
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub

Private Sub cb_Click(Index As Integer)
    On Error GoTo error
    Dim nSQL As String
    Dim xCampos() As String
    Dim nTitulo As String
    Dim xRs As New ADODB.Recordset
    
    If QueHace = 3 Then Exit Sub
    Select Case Index
    Case 0 '--MONEDA
    
        ReDim xCampos(2, 3) As String
        xCampos(0, 0) = "Moneda":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
        xCampos(1, 0) = "Símbolo":   xCampos(1, 1) = "simbolo":    xCampos(1, 2) = "1000":   xCampos(1, 3) = "C"
        nSQL = "SELECT mae_moneda.id, mae_moneda.descripcion AS nombre, mae_moneda.id AS cod, mae_moneda.simbolo FROM mae_moneda ORDER BY mae_moneda.descripcion;"
        nTitulo = "Buscar Moneda"
    Case 1 '--SUBLIBRO
        If NulosN(TxtIdLibro.Text) = 0 Then
            MsgBox "Falta especificar el libro", vbExclamation, xTitulo
            
            Exit Sub
        End If
        
        ReDim xCampos(2, 3) As String
        xCampos(0, 0) = "Descripción":   xCampos(0, 1) = "nombre":  xCampos(0, 2) = "5000":     xCampos(0, 3) = "C"
        xCampos(1, 0) = "Id":            xCampos(1, 1) = "id":      xCampos(1, 2) = "500":      xCampos(1, 3) = "N"
        
        nSQL = "SELECT mae_librossub.id, mae_librossub.descripcion AS nombre, mae_librossub.id AS cod " _
            + vbCr + " From mae_librossub " _
            + vbCr + " Where (((mae_librossub.idlib) = " & NulosN(TxtIdLibro.Text) & ")) " _
            + vbCr + " ORDER BY mae_librossub.descripcion;"

        nTitulo = "Buscar Sub Libro"
    End Select

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
    If xRs.State = 0 Then GoTo SALIR
    If xRs.BOF = True Or xRs.EOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    If xRs.State = 0 Then GoTo SALIR
    
    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cb_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
    Select Case Index
        Case 0 'MONEDA
             TxtFchEmi.SetFocus
        Case 1 'SUB LIBRO
            txt_cb(0).SetFocus
    End Select
   
SALIR:
    Set xRs = Nothing
Exit Sub

error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click (" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cb_cod(Index).Caption = ""
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

Sub CrearCabeceraVS(numPag As Integer)
    Dim xCad As String

    FrmVsPrinter.Vs.TextAlign = taLeftTop
    FrmVsPrinter.Vs.FontName = "Courier New"
    FrmVsPrinter.Vs.FontBold = True
    FrmVsPrinter.Vs.FontSize = 9

    FrmVsPrinter.Vs.CurrentX = 1000:      FrmVsPrinter.Vs.CurrentY = 1200
    FrmVsPrinter.Vs.Paragraph = "EMPRESA   : " & NomEmp

    FrmVsPrinter.Vs.CurrentX = 7950:      FrmVsPrinter.Vs.CurrentY = 1200
    FrmVsPrinter.Vs.Paragraph = "FECHA        : " & Format(Date, "dd/mm/yy")

    FrmVsPrinter.Vs.CurrentX = 1000:      FrmVsPrinter.Vs.CurrentY = 1400
    FrmVsPrinter.Vs.Paragraph = "Nº R.U.C. : " & NumRUC

    FrmVsPrinter.Vs.CurrentX = 7950:      FrmVsPrinter.Vs.CurrentY = 1400
    FrmVsPrinter.Vs.Paragraph = "Nº PÁGINA    : " & Format(numPag, "0000")

    FrmVsPrinter.Vs.DrawLine 1000, 1650, 11000, 1650
End Sub

Private Sub IMPRIMIR(Tipo As Integer, linea As Integer, Cabecera As Boolean)
'01/08/11 Modificado Johan Castro
'         Se agrega la impresion de los destinos de las cuentas, se elimina lineas de codigo cuando se imprime el detalle
'17/01/12 Modificado Johan Castro
'         1.- Cambiar campo en sentencia SQL. Antes mae_documento.abrev AS destipdoc ahora mae_documento.descripcion AS destipdoc
'         2.- Quitar declaracion de RstAux.
'         3.- Reemplazar codigo para mostrar el tipo de documento en la impresión. Antes .TextBox RstFrm("descripcion"),3000, xLinea, 2375, 250, True, False, False ahora .TextBox NulosC(xRs("destipdoc")),3000, xLinea, 2375, 250, True, False, False


'    tipo = 1: Imprimir listado
'    tipo = 2: Imprimir Unitario

'    cabecera = True: Imprimir con cabecera
'    cabecera = False: Imprimir sin cabecera

    Dim A As Integer
    Dim xLinea As Integer
    Dim xform As New eps_librerias.FormSeleccion
    Dim xRs As New ADODB.Recordset
    Dim nSQLFiltro As String '--Almacenara el filtro por movimiento
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(8, 5) As String
    Dim numeroPag As Integer
    Dim nSQL As String
    
    
    numeroPag = 1
    
''    Select Case Tipo
''        Case 1
    '--solo cuando se imprima el detalle se aplicara el filtro
    If Tipo = 2 Then nSQLFiltro = " and con_proviciones.id= " & RstFrm("id")

            'consulta para obtener listado de
    xform.SqlCad = "SELECT 0 as xsel, con_proviciones.*, mae_libros.descripcion AS desclib, mae_documento.descripcion AS destipdoc, con_meses.descripcion AS descmes, " _
        + vbCr + " (SELECT Sum([con_provicionesdet]![imp]) AS totdeb  FROM con_provicionesdet WHERE (((con_provicionesdet.id)=con_proviciones.id)  AND ((con_provicionesdet.tipo)=0))) & '' AS totdeb1, " _
        + vbCr + " (SELECT Sum([con_provicionesdet]![imp]) AS totdeb FROM con_provicionesdet WHERE  (((con_provicionesdet.id)=con_proviciones.id)   AND ((con_provicionesdet.tipo)=-1))) & '' AS tothab1, " _
        + vbCr + " mae_moneda.descripcion AS mondesc, mae_moneda.simbolo, [con_proviciones]![numser]+'-'+[con_proviciones]![numdoc] AS numedoc, mae_librossub.descripcion AS sublibdesc, " _
        + vbCr + " Format([con_proviciones].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & Mid([con_proviciones].[numreg],3) AS registro " _
        + vbCr + " FROM ((((con_proviciones LEFT JOIN mae_libros ON con_proviciones.idlib = mae_libros.id) LEFT JOIN con_meses ON con_proviciones.idmes = con_meses.id) LEFT JOIN mae_moneda ON con_proviciones.idmon = mae_moneda.id) LEFT JOIN mae_documento ON con_proviciones.tipdoc = mae_documento.id) LEFT JOIN mae_librossub ON con_proviciones.idsublib = mae_librossub.id " _
        + vbCr + " Where (((con_proviciones.ano) = " & AnoTra & "  ) And ((con_proviciones.idmes) = " & mMesActivo & "  )) " & nSQLFiltro _
        + vbCr + " ORDER BY con_proviciones.fchreg;"
    
    
    If Tipo = 1 Then
            xCampos(0, 0) = "Nº Registro":      xCampos(0, 1) = "registro":         xCampos(0, 2) = "950":      xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
            xCampos(1, 0) = "Nº Documento":     xCampos(1, 1) = "numedoc":          xCampos(1, 2) = "1500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
            xCampos(2, 0) = "M":                xCampos(2, 1) = "simbolo":          xCampos(2, 2) = "700":      xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
            xCampos(3, 0) = "Fecha":            xCampos(3, 1) = "fchreg":           xCampos(3, 2) = "1000":     xCampos(3, 3) = "D":    xCampos(3, 4) = "N"
            xCampos(4, 0) = "Sub Libro":        xCampos(4, 1) = "sublibdesc":       xCampos(4, 2) = "1000":     xCampos(4, 3) = "C":    xCampos(4, 4) = "N"
            xCampos(5, 0) = "Glosa":            xCampos(5, 1) = "glosa":            xCampos(5, 2) = "3200":     xCampos(5, 3) = "C":    xCampos(5, 4) = "N"
            xCampos(6, 0) = "Debe":             xCampos(6, 1) = "totdeb1":          xCampos(6, 2) = "1000":     xCampos(6, 3) = "N":    xCampos(6, 4) = "N"
            xCampos(7, 0) = "Haber":            xCampos(7, 1) = "tothab1":          xCampos(7, 2) = "1000":     xCampos(7, 3) = "N":    xCampos(7, 4) = "N"
                        
                
            xform.Titulo = "Operaciones a Imprimir"
            Set xform.Coneccion = xCon
            Set xRs = Nothing
            Set xRs = xform.Seleccionar(xCampos)
    Else
        RST_Busq xRs, xform.SqlCad, xCon
    End If
            
            If xRs.State = 1 Then
                If xRs.RecordCount <> 0 Then
                    xRs.MoveFirst
            
                    FrmVsPrinter.Vs.StartDoc
                    FrmVsPrinter.Vs.BrushColor = &H80000005
                    xLinea = 1100
                    
                    '-----Encabezado
                    CrearCabeceraVS numeroPag
                    For A = 1 To xRs.RecordCount
                        With FrmVsPrinter.Vs
                            xLinea = xLinea + 600
                            
                            If xLinea >= 14000 Then
                                .NewPage
                                numeroPag = numeroPag + 1
                                CrearCabeceraVS numeroPag
                                xLinea = 1700
                            End If
                            
                            .FontSize = 15
                            .TextAlign = taCenterMiddle
                            .TextBox "Provisiones Diversas", 1000, xLinea, 7500, 500, True, False, True
                                    
                            .FontSize = 10
                            .TextBox "N° Registro", 8600, xLinea, 2375, 250, True, False, True
                            .TextBox NulosC(xRs("registro")), 8600, xLinea + 250, 2375, 250, True, False, True
                            
                            '-----Descripcion
                            .TextAlign = taLeftMiddle
                            xLinea = xLinea + 500
                            .TextBox "Libro : ", 1000, xLinea, 2375, 250, True, False, False
                            .TextBox NulosC(xRs("desclib")), 3000, xLinea, 2375, 250, True, False, False
                            
                            .TextBox "Sub Libro : ", 7000, xLinea, 2375, 250, True, False, False
                            .TextBox NulosC(xRs("sublibdesc")), 9000, xLinea, 2375, 250, True, False, False
                            xLinea = xLinea + 250
                            
                            .TextBox "Fch.Doc : ", 1000, xLinea, 2375, 250, True, False, False
                            .TextBox NulosC(xRs("fchdoc")), 3000, xLinea, 2375, 250, True, False, False
                            
                            .TextBox "T.C. : ", 7000, xLinea, 2375, 250, True, False, False
                            .TextBox NulosN(xRs("tc")), 9000, xLinea, 2375, 250, True, False, False
                            xLinea = xLinea + 250

                            .TextBox "Tipo Documento : ", 1000, xLinea, 2375, 250, True, False, False
                            
                            .TextBox NulosC(xRs("destipdoc")), 3000, xLinea, 2375, 250, True, False, False
                            
                            .TextBox "N° Documento :", 7000, xLinea, 2375, 250, True, False, False
                            .TextBox NulosC(xRs("numser")) & "-" & NulosC(xRs("numdoc")), 9000, xLinea, 2375, 250, True, False, False
                            xLinea = xLinea + 250
                            
                            .TextBox "Moneda : ", 1000, xLinea, 2375, 250, True, False, False
                            .TextBox NulosC(xRs("mondesc")), 3000, xLinea, 2375, 250, True, False, False
                            xLinea = xLinea + 250
                            
                            .TextBox "Glosa : ", 1000, xLinea, 2375, 250, True, False, False
                            .TextBox NulosC(xRs("glosa")), 3000, xLinea, 8000, 250, True, False, False
                            xLinea = xLinea + 500
                            
                            If xLinea >= 15500 Then
                                .NewPage
                                numeroPag = numeroPag + 1
                                CrearCabeceraVS numeroPag
                                xLinea = 1700
                            End If
                            
                            .FontSize = 10
                            .TextAlign = taCenterMiddle
                            
                            '-----Contenido
                            .TextBox "ASIENTO CONTABLE", 1000, xLinea, 10000, 300, True, False, True
                            xLinea = xLinea + 350
                            
                            .TextBox "EXPRESADO EN MN", 5800, xLinea, 2600, 300, True, False, True
                            .TextBox "EXPRESADO EN ME", 8400, xLinea, 2600, 300, True, False, True
                            
                            .TextBox "N° Cuenta", 1000, xLinea, 800, 600, True, False, True
                            .TextBox "Descripción", 1800, xLinea, 4000, 600, True, False, True
                            
                            xLinea = xLinea + 300
                            .TextBox "Debe", 5800, xLinea, 1300, 300, True, False, True
                            .TextBox "Haber", 7100, xLinea, 1300, 300, True, False, True
                            .TextBox "Debe", 8400, xLinea, 1300, 300, True, False, True
                            .TextBox "Haber", 9700, xLinea, 1300, 300, True, False, True
                            
                    
                            xLinea = xLinea + 300
                            .FontSize = 7
                            .TextAlign = taLeftMiddle
                            
                            Dim RstDet As New ADODB.Recordset
                            '--Consulta con destinos detallados
'                            nSQL = "SELECT con_provicionesdet.id, con_provicionesdet.idcuen, con_planctas.cuenta, con_planctas.descripcion, IIf([con_provicionesdet].[tipo]=0,[con_provicionesdet]![imp],0) AS debe, IIf([con_provicionesdet].[tipo]=-1,[con_provicionesdet]![imp],0) AS haber, " _
                                & " (SELECT Count([idcuen]) AS numdoc From con_provicionesdetdoc WHERE (((con_provicionesdetdoc.idprov)=con_provicionesdet.id) AND ((con_provicionesdetdoc.idcuen)=con_provicionesdet.idcuen))) AS cantdoc, con_planctas.idmodulo " _
                                + vbCr + " FROM con_provicionesdet LEFT JOIN con_planctas ON con_provicionesdet.idcuen = con_planctas.id " _
                                + vbCr + " Where (((con_provicionesdet.id) = " & xRs("id") & ")) " _
                                + vbCr + "UNION ALL " _
                                + vbCr + " SELECT con_provicionesdet.id, con_planctas.ctadesdeb AS idcuen, con_planctas_1.cuenta, con_planctas_1.descripcion, con_provicionesdet.imp AS debe, 0 AS haber, 0 AS cantdoc, con_planctas.idmodulo " _
                                + vbCr + " FROM (con_provicionesdet LEFT JOIN con_planctas ON con_provicionesdet.idcuen = con_planctas.id) LEFT JOIN con_planctas AS con_planctas_1 ON con_planctas.ctadesdeb = con_planctas_1.id " _
                                + vbCr + " Where (((con_provicionesdet.id) = " & xRs("id") & ") And ((con_planctas.ctadesdeb) <> 0)) " _
                                + vbCr + " Union All " _
                                + vbCr + " SELECT con_provicionesdet.id, con_planctas.ctadeshab AS idcuen, con_planctas_1.cuenta, con_planctas_1.descripcion, 0 AS debe, con_provicionesdet.[imp] AS haber, 0 AS cantdoc, con_planctas.idmodulo " _
                                + vbCr + " FROM (con_provicionesdet LEFT JOIN con_planctas ON con_provicionesdet.idcuen = con_planctas.id) LEFT JOIN con_planctas AS con_planctas_1 ON con_planctas.ctadeshab = con_planctas_1.id " _
                                + vbCr + " Where (((con_provicionesdet.id) = " & xRs("id") & ") And ((con_planctas.ctadeshab) <> 0)) "
                            
                            '--Consulta con destinos resumidos
                            nSQL = "SELECT con_provicionesdet.id, con_provicionesdet.idcuen, con_planctas.cuenta, con_planctas.descripcion, IIf([con_provicionesdet].[tipo]=0,[con_provicionesdet]![imp],0) AS debe, IIf([con_provicionesdet].[tipo]=-1,[con_provicionesdet]![imp],0) AS haber, " _
                                & " (SELECT Count([idcuen]) AS numdoc From con_provicionesdetdoc WHERE (((con_provicionesdetdoc.idprov)=con_provicionesdet.id) AND ((con_provicionesdetdoc.idcuen)=con_provicionesdet.idcuen))) AS cantdoc, con_planctas.idmodulo " _
                                + vbCr + " FROM con_provicionesdet LEFT JOIN con_planctas ON con_provicionesdet.idcuen = con_planctas.id " _
                                + vbCr + " Where (((con_provicionesdet.id) = " & xRs("id") & ")) " _
                                + vbCr + " UNION ALL " _
                                + vbCr + " SELECT con_provicionesdet.id, con_planctas.ctadesdeb AS idcuen, con_planctas_1.cuenta, con_planctas_1.descripcion, Sum(con_provicionesdet.[imp]) AS debe, 0 AS haber, 0 AS cantdoc, con_planctas.idmodulo " _
                                + vbCr + " FROM (con_provicionesdet LEFT JOIN con_planctas ON con_provicionesdet.idcuen = con_planctas.id) LEFT JOIN con_planctas AS con_planctas_1 ON con_planctas.ctadesdeb = con_planctas_1.id " _
                                + vbCr + " GROUP BY con_provicionesdet.id, con_planctas.ctadesdeb, con_planctas_1.cuenta, con_planctas_1.descripcion, con_planctas.idmodulo " _
                                + vbCr + " HAVING (((con_provicionesdet.id)=" & xRs("id") & ") AND ((con_planctas.ctadesdeb)<>0)) " _
                                + vbCr + " UNION ALL " _
                                + vbCr + " SELECT con_provicionesdet.id, con_planctas.ctadeshab AS idcuen, con_planctas_1.cuenta, con_planctas_1.descripcion, 0 AS debe, Sum(con_provicionesdet.[imp]) AS haber, 0 AS cantdoc, con_planctas.idmodulo " _
                                + vbCr + " FROM (con_provicionesdet LEFT JOIN con_planctas ON con_provicionesdet.idcuen = con_planctas.id) LEFT JOIN con_planctas AS con_planctas_1 ON con_planctas.ctadeshab = con_planctas_1.id " _
                                + vbCr + " GROUP BY con_provicionesdet.id, con_planctas.ctadeshab, con_planctas_1.cuenta, con_planctas_1.descripcion,  con_planctas.idmodulo " _
                                + vbCr + " HAVING (((con_provicionesdet.id)=" & xRs("id") & ") AND ((con_planctas.ctadeshab)<>0))  "

                            RST_Busq RstDet, nSQL, xCon
                            
                            Agregando = True
                            Dim sumaDeb As Double
                            Dim sumaHab As Double
                            If RstDet.RecordCount <> 0 Then
                                RstDet.MoveFirst
                                '--aplicar orden
                                RstDet.Sort = "cuenta"
                                
                                sumaDeb = 0
                                sumaHab = 0
                                Dim B As Integer
                                For B = 1 To RstDet.RecordCount
                                    If xLinea >= 15500 Then
                                        .FontSize = 9
                                        .TextAlign = taRightMiddle
                                        .TextBox "VAN ", 1800, xLinea, 4000, 300, True, False, False
                                        .FontSize = 7
                                        'van debe MN
                                        .TextBox Format(NulosN(sumaDeb), FORMAT_MONTO), 5800, xLinea, 1300, 300, True, False, True
                                        'van haber MN
                                        .TextBox Format(NulosN(sumaHab), FORMAT_MONTO), 7100, xLinea, 1300, 300, True, False, True
                                        'van debe ME
                                        .TextBox Format(NulosN(sumaDeb) / NulosN(xRs("tc")), FORMAT_MONTO), 8400, xLinea, 1300, 300, True, False, True
                                        'van haber ME
                                        .TextBox Format(NulosN(sumaHab) / NulosN(xRs("tc")), FORMAT_MONTO), 9700, xLinea, 1300, 300, True, False, True
                                    
                                        .NewPage
                                        numeroPag = numeroPag + 1
                                        CrearCabeceraVS numeroPag
                                        xLinea = 1700
                                        
                                        .FontSize = 9
                                        .TextAlign = taRightMiddle
                                        .TextBox "VIENEN ", 1800, xLinea, 4000, 300, True, False, False
                                        .FontSize = 7
                                        'vienen debe MN
                                        .TextBox Format(NulosN(sumaDeb), FORMAT_MONTO), 5800, xLinea, 1300, 300, True, False, True
                                        'vienen haber MN
                                        .TextBox Format(NulosN(sumaHab), FORMAT_MONTO), 7100, xLinea, 1300, 300, True, False, True
                                        'vienen debe ME
                                        .TextBox Format(NulosN(sumaDeb) / NulosN(xRs("tc")), FORMAT_MONTO), 8400, xLinea, 1300, 300, True, False, True
                                        'vienen haber ME
                                        .TextBox Format(NulosN(sumaHab) / NulosN(xRs("tc")), FORMAT_MONTO), 9700, xLinea, 1300, 300, True, False, True
                                        
                                        xLinea = xLinea + 300
                                    End If
                                
                                
                                
                                    .TextAlign = taLeftMiddle
                                    .TextBox " " & NulosC(RstDet("cuenta")), 1000, xLinea, 800, 300, True, False, True
                                    .TextBox " " & NulosC(RstDet("descripcion")), 1800, xLinea, 4000, 300, True, False, True
    '
                                    .TextAlign = taRightMiddle
                                    'deben MN
                                    .TextBox Format(NulosN(RstDet("debe")), FORMAT_MONTO) & " ", 5800, xLinea, 1300, 300, True, False, True
                                    sumaDeb = sumaDeb + NulosN(RstDet("debe"))
                                    'haber MN
                                    .TextBox Format(NulosN(RstDet("haber")), FORMAT_MONTO) & " ", 7100, xLinea, 1300, 300, True, False, True
                                    sumaHab = sumaHab + NulosN(RstDet("haber"))
                                    
                                    'deben ME
                                    .TextBox Format(NulosN(RstDet("debe")) / NulosN(xRs("tc")), FORMAT_MONTO) & " ", 8400, xLinea, 1300, 300, True, False, True
                                    'haber ME
                                    .TextBox Format(NulosN(RstDet("haber")) / NulosN(xRs("tc")), FORMAT_MONTO) & " ", 9700, xLinea, 1300, 300, True, False, True
                                    xLinea = xLinea + 300
                                
                                    RstDet.MoveNext
                                    If RstDet.EOF = True Then Exit For
                                Next B
                            End If
                            
                            xLinea = xLinea + 100

                            .FontSize = 10
                            .TextAlign = taRightMiddle
                            .TextBox "Total ", 1800, xLinea, 4000, 300, True, False, True
'
                            .FontSize = 7
                            .TextBox Format(NulosN(sumaDeb), FORMAT_MONTO) & " ", 5800, xLinea, 1300, 300, True, False, True
                            .TextBox Format(NulosN(sumaHab), FORMAT_MONTO) & " ", 7100, xLinea, 1300, 300, True, False, True
                            
                            .TextBox Format(NulosN(sumaDeb) / NulosN(xRs("tc")), FORMAT_MONTO) & " ", 8400, xLinea, 1300, 300, True, False, True
                            .TextBox Format(NulosN(sumaHab) / NulosN(xRs("tc")), FORMAT_MONTO) & " ", 9700, xLinea, 1300, 300, True, False, True
                            
                        End With
                        xRs.MoveNext
                        If xRs.EOF = True Then Exit For
                    Next A
                    FrmVsPrinter.Vs.EndDoc
                End If
            Else
                Exit Sub
            End If
''        Case 2
''            numeroPag = 1
''            With FrmVsPrinter.Vs
''                .StartDoc
''                .BrushColor = &H80000005
''                xLinea = 1700
''
''                '-----Encabezado
''                CrearCabeceraVS numeroPag
''
''                .FontSize = 15
''                .TextAlign = taCenterMiddle
''                .TextBox "Provisiones Diversas", 1000, xLinea, 7500, 500, True, False, True
''
''                .FontSize = 10
''                .TextBox "N° Registro", 8600, xLinea, 2375, 250, True, False, True
''                .TextBox NulosC(RstFrm("registro")), 8600, xLinea + 250, 2375, 250, True, False, True
''
''                '-----Descripcion
''                .TextAlign = taLeftMiddle
''                xLinea = xLinea + 500
''                .TextBox "Libro : ", 1000, xLinea, 2375, 250, True, False, False
''                .TextBox LblLibro.Caption, 3000, xLinea, 2375, 250, True, False, False
''
''                .TextBox "Sub Libro : ", 7000, xLinea, 2375, 250, True, False, False
''                .TextBox lbl_cb(1).Caption, 9000, xLinea, 2375, 250, True, False, False
''                xLinea = xLinea + 250
''
''                .TextBox "Fch.Doc : ", 1000, xLinea, 2375, 250, True, False, False
''                .TextBox TxtFchEmi.Valor, 3000, xLinea, 2375, 250, True, False, False
''
''                .TextBox "T.C. : ", 7000, xLinea, 2375, 250, True, False, False
''                .TextBox NulosN(TxtTC.Text), 9000, xLinea, 2375, 250, True, False, False
''                xLinea = xLinea + 250
''
''                .TextBox "Tipo Documento : ", 1000, xLinea, 2375, 250, True, False, False
''                .TextBox LblDocumento.Caption, 3000, xLinea, 2375, 250, True, False, False
''
''                .TextBox "N° Documento :", 7000, xLinea, 2375, 250, True, False, False
''                .TextBox TxtSerDoc.Text & "-" & TxtNumDoc.Text, 9000, xLinea, 2375, 250, True, False, False
''                xLinea = xLinea + 250
''
''                .TextBox "Moneda : ", 1000, xLinea, 2375, 250, True, False, False
''                .TextBox lbl_cb(0).Caption, 3000, xLinea, 2375, 250, True, False, False
''                xLinea = xLinea + 250
''
''                .TextBox "Glosa : ", 1000, xLinea, 2375, 250, True, False, False
''                .TextBox Trim(TxtGlosa.Text), 3000, xLinea, 8000, 250, True, False, False
''                xLinea = xLinea + 500
''
''                .FontSize = 10
''                .TextAlign = taCenterMiddle
''
''                '-----Contenido
''                .TextBox "ASIENTO CONTABLE", 1000, xLinea, 10000, 300, True, False, True
''                xLinea = xLinea + 350
''
''                .TextBox "EXPRESADO EN MN", 5800, xLinea, 2600, 300, True, False, True
''                .TextBox "EXPRESADO EN ME", 8400, xLinea, 2600, 300, True, False, True
''
''                .TextBox "N° Cta.", 1000, xLinea, 800, 600, True, False, True
''                .TextBox "Descripción", 1800, xLinea, 4000, 600, True, False, True
''
''                xLinea = xLinea + 300
''                .TextBox "Debe", 5800, xLinea, 1300, 300, True, False, True
''                .TextBox "Haber", 7100, xLinea, 1300, 300, True, False, True
''                .TextBox "Debe", 8400, xLinea, 1300, 300, True, False, True
''                .TextBox "Haber", 9700, xLinea, 1300, 300, True, False, True
''
''
''                xLinea = xLinea + 300
''                .FontSize = 7
''                .TextAlign = taLeftMiddle
''                Dim debeAux As Double
''                Dim haberAux As Double
''                For A = 1 To fg1.Rows - 1
''                    If xLinea >= 15500 Then
''                        .FontSize = 8
''                        .TextAlign = taRightMiddle
''                        .TextBox "VAN ", 1800, xLinea, 4000, 300, True, False, False
''                        .FontSize = 7
''                        'van debe
''                        .TextBox Format(NulosN(debeAux), FORMAT_MONTO), 5800, xLinea, 1300, 300, True, False, True
''                        .TextBox Format(NulosN(debeAux) / NulosN(TxtTC.Text), FORMAT_MONTO), 7100, xLinea, 1300, 300, True, False, True
''                        'van haber
''                        .TextBox Format(NulosN(haberAux), FORMAT_MONTO), 8400, xLinea, 1300, 300, True, False, True
''                        .TextBox Format(NulosN(haberAux) / NulosN(TxtTC.Text), FORMAT_MONTO), 9700, xLinea, 1300, 300, True, False, True
''
''                        .NewPage
''                        numeroPag = numeroPag + 1
''                        CrearCabeceraVS numeroPag
''                        xLinea = 1700
''
''                        .FontSize = 8
''                        .TextAlign = taRightMiddle
''                        .TextBox "VIENEN ", 1800, xLinea, 4000, 300, True, False, False
''                        .FontSize = 7
''                        'vienen debe
''                        .TextBox Format(NulosN(debeAux), FORMAT_MONTO), 5800, xLinea, 1300, 300, True, False, True
''                        .TextBox Format(NulosN(debeAux) / NulosN(TxtTC.Text), FORMAT_MONTO), 7100, xLinea, 1300, 300, True, False, True
''                        'vienen haber
''                        .TextBox Format(NulosN(haberAux), FORMAT_MONTO), 8400, xLinea, 1300, 300, True, False, True
''                        .TextBox Format(NulosN(haberAux) / NulosN(TxtTC.Text), FORMAT_MONTO), 9700, xLinea, 1300, 300, True, False, True
''
''                        xLinea = xLinea + 300
''                    End If
''
''                    .TextAlign = taLeftMiddle
''                    'numero de cuenta
''                    .TextBox " " & fg1.TextMatrix(A, 1), 1000, xLinea, 800, 300, True, False, True
''                    'descripcion
''                    .TextBox " " & fg1.TextMatrix(A, 2), 1800, xLinea, 4000, 300, True, False, True
''
''                    .TextAlign = taRightMiddle
''                    .TextBox Format(NulosN(fg1.TextMatrix(A, 3)), FORMAT_MONTO) & " ", 5800, xLinea, 1300, 300, True, False, True
''                    debeAux = debeAux + NulosN(fg1.TextMatrix(A, 3))
''                    .TextBox Format(NulosN(fg1.TextMatrix(A, 4)), FORMAT_MONTO) & " ", 7100, xLinea, 1300, 300, True, False, True
''                    haberAux = haberAux + NulosN(fg1.TextMatrix(A, 4))
''
''                    .TextBox Format(NulosN(fg1.TextMatrix(A, 3)) / NulosN(TxtTC.Text), FORMAT_MONTO) & " ", 8400, xLinea, 1300, 300, True, False, True
''                    .TextBox Format(NulosN(fg1.TextMatrix(A, 4)) / NulosN(TxtTC.Text), FORMAT_MONTO) & " ", 9700, xLinea, 1300, 300, True, False, True
''                    xLinea = xLinea + 300
''                Next A
''                xLinea = xLinea + 100
''
''                .FontSize = 10
''                .TextAlign = taRightMiddle
''                .TextBox "Total ", 1800, xLinea, 4000, 300, True, False, True
''
''                .FontSize = 7
''                .TextBox Format(NulosN(TxtTotDeb.Text), FORMAT_MONTO) & " ", 5800, xLinea, 1300, 300, True, False, True
''                .TextBox Format(NulosN(TxtTotHab.Text), FORMAT_MONTO) & " ", 7100, xLinea, 1300, 300, True, False, True
''
''                .TextBox Format(NulosN(TxtTotDeb.Text) / NulosN(TxtTC.Text), FORMAT_MONTO) & " ", 8400, xLinea, 1300, 300, True, False, True
''                .TextBox Format(NulosN(TxtTotHab.Text) / NulosN(TxtTC.Text), FORMAT_MONTO) & " ", 9700, xLinea, 1300, 300, True, False, True
''
''                .EndDoc
''            End With
''    End Select
    'se vuelve a cargar el Dg1
    OpcionesPeriodo
    'Se muestra el diseño de impresion
    FrmVsPrinter.Show
    
End Sub

Private Sub pExportarMSExcel()
    Dim A&, B&
    Dim xFilas&
    On Error GoTo error
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    'determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1
    'abre el Libro
    objExcel.WindowState = 1
    objExcel.Workbooks.Add
    With objExcel.ActiveSheet
        
        .Cells(1, 2) = NomEmp
        .Cells(1, 5) = Date
        .Cells(2, 2) = "Nº R.U.C. : " + NumRUC
        
        .Columns(2).ColumnWidth = Fg1.ColWidth(1) / 100
        .Columns(3).ColumnWidth = Fg1.ColWidth(2) / 100
        .Columns(4).ColumnWidth = Fg1.ColWidth(3) / 100
        .Columns(5).ColumnWidth = Fg1.ColWidth(4) / 100
                        
        '-----encabezado
        xFilas = 4
        .Cells(xFilas, 2) = "Provisiones Diversas"
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Libro"
        .Cells(xFilas, 3) = LblLibro.Caption
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Sub Libro"
        .Cells(xFilas, 3) = lbl_cb(1).Caption
        .Cells(xFilas, 4) = "Periodo"
        .Cells(xFilas, 5) = lblperiodo(1).Caption
        
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Moneda"
        .Cells(xFilas, 3) = lbl_cb(0).Caption
        
        .Cells(xFilas, 4) = "T.C."
        .Cells(xFilas, 5) = NulosN(TxtTC.Text)
        
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Fch.Doc"
        .Cells(xFilas, 3) = "'" & TxtFchEmi.Valor
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Tipo Documento"
        .Cells(xFilas, 3) = "'" & LblDocumento.Caption
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "N° Documento"
        .Cells(xFilas, 3) = "'" & TxtSerDoc.Text & "-" & TxtNumDoc.Text
        xFilas = xFilas + 1
        .Cells(xFilas, 2) = "Glosa"
        .Cells(xFilas, 3) = "'" & Trim(TxtGlosa.Text)

        xFilas = xFilas + 2
        
        .Cells(xFilas, 2) = "N° Cuenta"
        .Cells(xFilas, 3) = "Descripción"
        .Cells(xFilas, 4) = "Debe"
        .Cells(xFilas, 5) = "Haber"
        
        xFilas = xFilas + 1
        For A = 1 To Fg1.Rows - 1
            .Cells(xFilas, 2) = "'" + Fg1.TextMatrix(A, 1)
            .Cells(xFilas, 3) = "'" + Fg1.TextMatrix(A, 2)
            .Cells(xFilas, 4) = NulosN(Fg1.TextMatrix(A, 3))
            .Cells(xFilas, 5) = NulosN(Fg1.TextMatrix(A, 4))
            xFilas = xFilas + 1
        Next A

        .Cells(xFilas, 3) = "Total"
        .Cells(xFilas, 4) = NulosN(TxtTotDeb.Text)
        .Cells(xFilas, 5) = NulosN(TxtTotHab.Text)
        
    End With
    
    MsgBox "El Registro se exportó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    objExcel.WindowState = 1
    objExcel.Visible = True
    Set objExcel = Nothing
    Exit Sub
error:
    Set objExcel = Nothing
    SHOW_ERROR Me.Name, "ExportarExcelDetalle", , IIf(Err.Number <> 50290, "", "No manipule el archivo hasta que termine de exportar!!!!")
End Sub

Sub PreparaRST_Tmp()
    'Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(21, 3) As String

    xCampos(0, 0) = "fchdoc":     xCampos(0, 1) = "C":      xCampos(0, 2) = "10"  'fecha del documento
    xCampos(1, 0) = "numdoc":     xCampos(1, 1) = "C":      xCampos(1, 2) = "15" 'numero del documento
    xCampos(2, 0) = "nommon":     xCampos(2, 1) = "C":      xCampos(2, 2) = "20" 'descrpcion de la moneda
    xCampos(3, 0) = "provee":     xCampos(3, 1) = "C":      xCampos(3, 2) = "100"  'nombre del proveedor
    xCampos(4, 0) = "nomcon":     xCampos(4, 1) = "C":      xCampos(4, 2) = "20" 'descripcion de la condicion de venta
    xCampos(5, 0) = "fchven":     xCampos(5, 1) = "F":      xCampos(5, 2) = "10" 'fecha de vencimiento
    xCampos(6, 0) = "impbru":     xCampos(6, 1) = "D":      xCampos(6, 2) = "20" 'importe bruto
    xCampos(7, 0) = "impigv":     xCampos(7, 1) = "D":      xCampos(7, 2) = "20" 'importe del igv
    xCampos(8, 0) = "impisc":     xCampos(8, 1) = "D":      xCampos(8, 2) = "20" 'importe del isc
    xCampos(9, 0) = "imptot":     xCampos(9, 1) = "D":      xCampos(9, 2) = "20" ' importe total del documento
    xCampos(10, 0) = "impsal":    xCampos(10, 1) = "D":     xCampos(10, 2) = "20" 'importe del saldo del documento
    xCampos(11, 0) = "iddoc":     xCampos(11, 1) = "N":     xCampos(11, 2) = "2" 'id del documento que se esta cargando
    xCampos(12, 0) = "idmon":     xCampos(12, 1) = "N":     xCampos(12, 2) = "2" 'id de la moneda del documento
    xCampos(13, 0) = "idcon":     xCampos(13, 1) = "N":     xCampos(13, 2) = "2" 'id de la condicion del documento
    xCampos(14, 0) = "idpro":     xCampos(14, 1) = "N":     xCampos(14, 2) = "2" 'id del proveedor
    xCampos(15, 0) = "idprovi":   xCampos(15, 1) = "N":     xCampos(15, 2) = "2" 'id de la condicion del documento
    xCampos(16, 0) = "idcuent":   xCampos(16, 1) = "N":     xCampos(16, 2) = "2" 'id del proveedor
    xCampos(17, 0) = "nuevo":     xCampos(17, 1) = "N":     xCampos(17, 2) = "2" 'especifica si es un documento nuevo
    xCampos(18, 0) = "edita":     xCampos(18, 1) = "N":     xCampos(18, 2) = "2" 'especifica si el documento se  puede modificar
    xCampos(19, 0) = "tipdoc":    xCampos(19, 1) = "N":     xCampos(19, 2) = "2" 'especifica si el documento se  puede modificar
    xCampos(20, 0) = "desdoc":    xCampos(20, 1) = "C":     xCampos(20, 2) = "20" 'especifica si el documento se  puede modificar
    Set RstTmp = xFun.CrearRstTMP(xCampos)

    RstTmp.Open
End Sub
'''
'''
'''
'''Function Grabarcopia() As Boolean
'''    If fValidarDatos() = False Then Exit Function
'''    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modficar") + " la Provición", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo) = vbNo Then Exit Function
'''
'''    Dim RstCab As New ADODB.Recordset
'''    Dim RstDet As New ADODB.Recordset
'''    Dim RstDetDoc As New ADODB.Recordset
'''    Dim RstDia As New ADODB.Recordset
'''    Dim xNumAsiento As String
'''    Dim xId2, A, B As Integer
'''
'''    On Error GoTo LaCague
'''    Me.MousePointer = vbHourglass
'''
'''    xCon.BeginTrans
'''    If QueHace = 1 Then
'''        xNumAsiento = NuevoNumAsiento(3, mMesActivo, xCon)
'''        xId2 = HallaCodigoTabla("con_proviciones", xCon, "id")
'''        RST_Busq RstCab, "SELECT TOP 1 * FROM con_proviciones", xCon
'''        RstCab.AddNew
'''        RstCab("id") = xId2
'''    Else
'''        xId2 = RstFrm("id")
'''        RST_Busq RstCab, "SELECT * FROM con_proviciones WHERE id = " & xId & "", xCon
'''
'''        'ELIMINAMOS EL DETALLE DE LA PROVICION
'''        xCon.Execute "DELETE * FROM con_provicionesdetdoc WHERE idprov = " & xId & ""
'''        xCon.Execute "DELETE * FROM con_provicionesdet WHERE id = " & xId & ""
'''
'''        xNumAsiento = DevuelveNumAsiento(3, RstFrm("id"), mMesActivo, xCon)
'''        If xNumAsiento = "" Then xNumAsiento = NuevoNumAsiento(3, mMesActivo, xCon)
'''        'ELIMINAMOS EL ASIENTO REGISTRADO EN EL DIARIO
'''        xCon.Execute "DELETE * FROM con_diario WHERE ((idmes = " & mMesActivo & ") and (idlib = 3) AND (idmov = " & xId & ")) ;"
'''    End If
'''
'''    RST_Busq RstDet, "SELECT TOP 1 * FROM con_provicionesdet", xCon
'''    RST_Busq RstDetDoc, "SELECT TOP 1 * FROM con_provicionesdetdoc", xCon
'''    RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
'''
'''    RstCab("ano") = AnoTra
'''    RstCab("idmes") = mMesActivo
'''    RstCab("numreg") = Format(mMesActivo, "00") + xNumAsiento
'''    If mMesActivo <> 0 And mMesActivo <> 13 Then
'''        RstCab("fchreg") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''    End If
'''    RstCab("idlib") = 3 '--proviciones diversas (libro diario)
'''    RstCab("idsublib") = NulosN(lbl_cb_cod(1).Caption)
'''    RstCab("idmon") = NulosN(lbl_cb_cod(0).Caption)
'''    RstCab("fchdoc") = CDate(TxtFchEmi.Valor)
'''    RstCab("tipdoc") = NulosN(TxtTipDoc.Text)
'''    RstCab("numser") = NulosC(TxtSerDoc.Text)
'''    RstCab("numdoc") = NulosC(TxtNumDoc.Text)
'''    RstCab("imp") = NulosN(TxtTotDeb.Text)
'''    RstCab("glosa") = NulosC(TxtGlosa.Text)
'''
'''    RstCab.Update
'''
'''    For A = 1 To Fg1.Rows - 1
'''        RstDet.AddNew
'''        RstDet("id") = xId2
'''        RstDet("idcuen") = NulosN(Fg1.TextMatrix(A, 5))
'''
'''        If NulosN(Fg1.TextMatrix(A, 3)) <> 0 Then
'''            RstDet("tipo") = 0 '--debe
'''            RstDet("imp") = NulosN(Fg1.TextMatrix(A, 3))
'''        End If
'''
'''        If NulosN(Fg1.TextMatrix(A, 4)) <> 0 Then
'''            RstDet("tipo") = -1 '--haber
'''            RstDet("imp") = NulosN(Fg1.TextMatrix(A, 4))
'''        End If
'''
'''        RstDet.Update
'''        'si la cuenta tiene detalle
'''        RstTmp.Filter = adFilterNone
'''        If RstTmp.RecordCount <> 0 Then
'''            RstTmp.MoveFirst
'''
'''            'seleccionamos los documentos antiguos para buscar el id en la tabla de documentos que corresponde
'''            RstTmp.Filter = "idprovi = " & xId & " AND idcuent = " & NulosN(Fg1.TextMatrix(A, 5)) & " AND nuevo = 0"
'''            If RstTmp.RecordCount <> 0 Then
'''                RstTmp.MoveFirst
'''                For B = 1 To RstTmp.RecordCount
'''
'''                    RstDetDoc.AddNew
'''                    RstDetDoc("idprov") = xId2
'''                    RstDetDoc("idcuen") = NulosN(Fg1.TextMatrix(A, 5))
'''                    RstDetDoc("idmod") = NulosN(Fg1.TextMatrix(A, 9))
'''                    RstDetDoc("iddoc") = RstTmp("iddoc")
'''
'''                    RstDetDoc.Update
'''                    RstTmp.MoveNext
'''                    If RstTmp.EOF = True Then Exit For
'''                Next B
'''            End If
'''        End If
'''    Next A
'''
'''    'escribimos los documentos de la provicion en su respectivo modulo
'''    Dim Rst As New ADODB.Recordset
'''    Dim xIdDoc As Integer
'''    Dim RstGraDocMod As New ADODB.Recordset 'recorset para grabar los documentos es su respectivo modulo
'''
'''    If QueHace = 1 Then  'si es nuevo asiento
'''        For A = 1 To Fg1.Rows - 1
'''            If NulosN(Fg1.TextMatrix(A, 8)) = -1 Then
'''                RstTmp.Filter = adFilterNone
'''                RstTmp.MoveFirst
'''                RstTmp.Filter = "idprovi = " & xId & " AND  idcuent  = " & Fg1.TextMatrix(A, 5) & ""
'''
'''                If NulosN(Fg1.TextMatrix(A, 6)) = 1 Then ' si es compras
'''                    RST_Busq RstGraDocMod, "SELECT TOP 1 * FROM com_compras", xCon
'''
'''                    RstTmp.MoveFirst
'''                    xIdDoc = HallaCodigoTabla("com_compras", xCon, "id")
'''
'''                    For B = 1 To RstTmp.RecordCount
'''                        RstGraDocMod.AddNew
'''                        RstGraDocMod("id") = xIdDoc
'''                        RstGraDocMod("idlib") = 1
'''                        RstGraDocMod("idtipo") = 5
'''                        RstGraDocMod("tipdoc") = RstTmp("tipdoc")
'''                        RstGraDocMod("idpro") = RstTmp("idpro")
'''                        RstGraDocMod("numser") = Mid(RstTmp("numdoc"), 1, 4)
'''                        RstGraDocMod("numdoc") = Mid(RstTmp("numdoc"), 6, 10)
'''                        RstGraDocMod("fchreg") = CDate("01/01/08")
'''                        RstGraDocMod("fchdoc") = RstTmp("fchdoc")
'''                        RstGraDocMod("fchven") = RstTmp("fchven")
'''                        RstGraDocMod("idconpag") = RstTmp("idcon")
'''                        RstGraDocMod("idmon") = RstTmp("idmon")
'''                        RstGraDocMod("impbru") = RstTmp("impbru")
'''                        RstGraDocMod("impina") = 0
'''                        RstGraDocMod("impisc") = RstTmp("impisc")
'''                        RstGraDocMod("impigv") = RstTmp("impigv")
'''                        RstGraDocMod("imptot") = RstTmp("imptot")
'''                        RstGraDocMod("impsal") = RstTmp("impsal")
'''                        RstGraDocMod("numreg") = Format(mMesActivo, "00") + xNumAsiento
'''
'''                        RstGraDocMod.Update
'''
'''                        RstDetDoc.AddNew
'''                        RstDetDoc("idprov") = xId2
'''                        RstDetDoc("idcuen") = Fg1.TextMatrix(A, 5)
'''                        RstDetDoc("idmod") = 1
'''                        RstDetDoc("iddoc") = xIdDoc
'''
'''                        RstDetDoc.Update
'''
'''                        RstTmp.MoveNext
'''                        If RstTmp.EOF = True Then Exit For
'''                        xIdDoc = xIdDoc + 1
'''                    Next B
'''                End If
'''
'''                If NulosN(Fg1.TextMatrix(A, 6)) = 2 Then ' si es ventas
'''                    RST_Busq RstGraDocMod, "SELECT TOP 1 * FROM vta_ventas", xCon
'''
'''                    RstTmp.MoveFirst
'''                    xIdDoc = HallaCodigoTabla("vta_ventas", xCon, "id")
'''
'''                    For B = 1 To RstTmp.RecordCount
'''                        RstGraDocMod.AddNew
'''                        RstGraDocMod("id") = xIdDoc
'''                        RstGraDocMod("idlib") = 2
'''                        RstGraDocMod("idtipo") = 2
'''                        RstGraDocMod("idcli") = RstTmp("idpro")
'''                        RstGraDocMod("idpunvencli") = 0
'''                        RstGraDocMod("tipdoc") = RstTmp("tipdoc")
'''                        RstGraDocMod("numser") = Mid(RstTmp("numdoc"), 1, 4)
'''                        RstGraDocMod("numdoc") = Mid(RstTmp("numdoc"), 6, 10)
'''                        RstGraDocMod("fchreg") = CDate("01/01/07")
'''                        RstGraDocMod("fchdoc") = RstTmp("fchdoc")
'''                        RstGraDocMod("fchven") = RstTmp("fchven")
'''                        RstGraDocMod("idconpag") = RstTmp("idcon")
'''                        RstGraDocMod("idmon") = RstTmp("idmon")
'''                        RstGraDocMod("impbru") = RstTmp("impbru")
'''                        RstGraDocMod("impinaf") = 0
'''                        RstGraDocMod("impisc") = RstTmp("impisc")
'''                        RstGraDocMod("impigv") = RstTmp("impigv")
'''                        RstGraDocMod("imptotdoc") = RstTmp("imptot")
'''                        RstGraDocMod("impsal") = RstTmp("imptot")
'''
'''                        RstGraDocMod("numreg") = Format(mMesActivo, "00") + xNumAsiento
'''                        RstGraDocMod("tipgen") = 1
'''                        RstGraDocMod.Update
'''
'''                        RstDetDoc.AddNew
'''                        RstDetDoc("idprov") = xId2
'''                        RstDetDoc("idcuen") = Fg1.TextMatrix(A, 5)
'''                        RstDetDoc("idmod") = 2
'''                        RstDetDoc("iddoc") = xIdDoc
'''
'''                        RstDetDoc.Update
'''
'''                        RstTmp.MoveNext
'''                        If RstTmp.EOF = True Then Exit For
'''                        xIdDoc = xIdDoc + 1
'''
'''                    Next B
'''                End If
'''
'''                If NulosN(Fg1.TextMatrix(A, 6)) = 4 Then ' si es ventas
'''                End If
'''
'''                If NulosN(Fg1.TextMatrix(A, 6)) = 8 Then ' si es ventas
'''                End If
'''            End If
'''        Next A
'''    End If
'''
'''
''''    xCampos(0, 0) = "fchdoc":     xCampos(0, 1) = "C":      xCampos(0, 2) = "10"  'fecha del documento
''''    xCampos(1, 0) = "numdoc":     xCampos(1, 1) = "C":      xCampos(1, 2) = "15" 'numero del documento
''''    xCampos(2, 0) = "nommon":     xCampos(2, 1) = "C":      xCampos(2, 2) = "20" 'descrpcion de la moneda
''''    xCampos(3, 0) = "provee":     xCampos(3, 1) = "C":      xCampos(3, 2) = "100"  'nombre del proveedor
''''    xCampos(4, 0) = "nomcon":     xCampos(4, 1) = "C":      xCampos(4, 2) = "20" 'descripcion de la condicion de venta
''''    xCampos(5, 0) = "fchven":     xCampos(5, 1) = "F":      xCampos(5, 2) = "10" 'fecha de vencimiento
''''    xCampos(6, 0) = "impbru":     xCampos(6, 1) = "D":      xCampos(6, 2) = "20" 'importe bruto
''''    xCampos(7, 0) = "impigv":     xCampos(7, 1) = "D":      xCampos(7, 2) = "20" 'importe del igv
''''    xCampos(8, 0) = "impisc":     xCampos(8, 1) = "D":      xCampos(8, 2) = "20" 'importe del isc
''''    xCampos(9, 0) = "imptot":     xCampos(9, 1) = "D":      xCampos(9, 2) = "20" ' importe total del documento
''''    xCampos(10, 0) = "impsal":    xCampos(10, 1) = "D":     xCampos(10, 2) = "20" 'importe del saldo del documento
''''    xCampos(11, 0) = "iddoc":     xCampos(11, 1) = "N":     xCampos(11, 2) = "2" 'id del documento que se esta cargando
''''    xCampos(12, 0) = "idmon":     xCampos(12, 1) = "N":     xCampos(12, 2) = "2" 'id de la moneda del documento
''''    xCampos(13, 0) = "idcon":     xCampos(13, 1) = "N":     xCampos(13, 2) = "2" 'id de la condicion del documento
''''    xCampos(14, 0) = "idpro":     xCampos(14, 1) = "N":     xCampos(14, 2) = "2" 'id del proveedor
''''    xCampos(15, 0) = "idprovi":   xCampos(15, 1) = "N":     xCampos(15, 2) = "2" 'id de la condicion del documento
''''    xCampos(16, 0) = "idcuent":   xCampos(16, 1) = "N":     xCampos(16, 2) = "2" 'id del proveedor
''''    xCampos(17, 0) = "nuevo":     xCampos(17, 1) = "N":     xCampos(17, 2) = "2" 'especifica si es un documento nuevo
''''    xCampos(18, 0) = "edita":     xCampos(18, 1) = "N":     xCampos(18, 2) = "2" 'especifica si el documento se  puede modificar
''''    xCampos(19, 0) = "tipdoc":    xCampos(19, 1) = "N":     xCampos(19, 2) = "2" 'especifica si el documento se  puede modificar
'''
'''    If QueHace = 2 Then  'si se modifica un asiento
'''    End If
'''
''''    For A = 1 To Fg1.Rows - 1
''''        If NulosN(Fg1.TextMatrix(A, 6)) = 1 Then ' si es compras
''''            'cargamos los documentos de la cuenta actual
''''            RstTmp.Filter = adFilterNone
''''            RstTmp.MoveFirst
''''            RstTmp.Filter = "idprovi = " & xId & " AND  idcuent  = " & Fg1.TextMatrix(A, 5) & ""
''''
''''            If RstTmp.RecordCount <> 0 Then
''''                'eliminamos lo documentos de com_compras que no esten en el rsttmp
''''                RST_Busq Rst, "SELECT com_compras.* From com_compras WHERE (((Mid([numreg],1,2))='" & Format(mMesActivo, "00") & "'))", xCon
''''                If Rst.RecordCount <> 0 Then
''''                    Rst.MoveFirst
''''                    For B = 1 To Rst.RecordCount
''''                        RstTmp.MoveFirst
''''                        RstTmp.Find "iddoc = " & Rst("id") & ""
''''                        If RstTmp.EOF = True Then
''''                            'no existe el documento en el rst RSTTMP, lo borramos de la tabla com_compras
''''                            xCon.Execute "DELETE * FROM com_compras WHERE id = " & Rst("id") & ""
''''                        End If
''''
''''                        Rst.MoveNext
''''                        If Rst.EOF = True Then Exit For
''''                    Next B
''''                    Rst.MoveFirst
''''                End If
''''            End If
''''
''''            'cargamos los documentos de la cuenta actual, solo se cargan los docuentos que se pueden editar
''''            RstTmp.Filter = adFilterNone
''''            RstTmp.MoveFirst
''''            RstTmp.Filter = "idprovi = " & xId & " AND  idcuent  = " & Fg1.TextMatrix(A, 5) & " And edita = 1"
''''
''''            If RstTmp.RecordCount <> 0 Then
''''                RstTmp.MoveFirst
''''                For B = 1 To RstTmp.RecordCount
''''                    'borramos los documentos de la tablas con_compras que esten en el RSTTMP
''''                    xCon.Execute "DELETE * FROM com_compras WHERE id = " & RstTmp("idddoc") & ""
''''                    RstTmp.MoveNext
''''                    If RstTmp.EOF = True Then Exit For
''''                Next B
''''            End If
''''        End If
''''    Next A
''''
'''''    xCampos(0, 0) = "fchdoc":     xCampos(0, 1) = "C":      xCampos(0, 2) = "10"  'fecha del documento
'''''    xCampos(1, 0) = "numdoc":     xCampos(1, 1) = "C":      xCampos(1, 2) = "15" 'numero del documento
'''''    xCampos(2, 0) = "nommon":     xCampos(2, 1) = "C":      xCampos(2, 2) = "20" 'descrpcion de la moneda
'''''    xCampos(3, 0) = "provee":     xCampos(3, 1) = "C":      xCampos(3, 2) = "100"  'nombre del proveedor
'''''    xCampos(4, 0) = "nomcon":     xCampos(4, 1) = "C":      xCampos(4, 2) = "20" 'descripcion de la condicion de venta
'''''    xCampos(5, 0) = "fchven":     xCampos(5, 1) = "F":      xCampos(5, 2) = "10" 'fecha de vencimiento
'''''    xCampos(6, 0) = "impbru":     xCampos(6, 1) = "D":      xCampos(6, 2) = "20" 'importe bruto
'''''    xCampos(7, 0) = "impigv":     xCampos(7, 1) = "D":      xCampos(7, 2) = "20" 'importe del igv
'''''    xCampos(8, 0) = "impisc":     xCampos(8, 1) = "D":      xCampos(8, 2) = "20" 'importe del isc
'''''    xCampos(9, 0) = "imptot":     xCampos(9, 1) = "D":      xCampos(9, 2) = "20" ' importe total del documento
'''''    xCampos(10, 0) = "impsal":    xCampos(10, 1) = "D":     xCampos(10, 2) = "20" 'importe del saldo del documento
'''''    xCampos(11, 0) = "iddoc":     xCampos(11, 1) = "N":     xCampos(11, 2) = "2" 'id del documento que se esta cargando
'''''    xCampos(12, 0) = "idmon":     xCampos(12, 1) = "N":     xCampos(12, 2) = "2" 'id de la moneda del documento
'''''    xCampos(13, 0) = "idcon":     xCampos(13, 1) = "N":     xCampos(13, 2) = "2" 'id de la condicion del documento
'''''    xCampos(14, 0) = "idpro":     xCampos(14, 1) = "N":     xCampos(14, 2) = "2" 'id del proveedor
'''''    xCampos(15, 0) = "idprovi":   xCampos(15, 1) = "N":     xCampos(15, 2) = "2" 'id de la condicion del documento
'''''    xCampos(16, 0) = "idcuent":   xCampos(16, 1) = "N":     xCampos(16, 2) = "2" 'id del proveedor
'''''    xCampos(17, 0) = "nuevo":     xCampos(17, 1) = "N":     xCampos(17, 2) = "2" 'especifica si es un documento nuevo
'''''    Set RstTmp = xFun.CrearRstTMP(xCampos)
'''
'''    'grabamos el diario
'''    For A = 1 To Fg1.Rows - 1
'''        RstDia.AddNew
'''        RstDia("año") = AnoTra
'''        RstDia("idmes") = mMesActivo
'''        RstDia("idlib") = 3 'NulosN(TxtIdLibro.Text)
'''        RstDia("idmov") = xId2
'''        RstDia("idcue") = NulosN(Fg1.TextMatrix(A, 5))
'''        RstDia("numasi") = xNumAsiento
'''        RstDia("tc") = NulosN(LblTipoCambio.Caption)
'''
'''        RstDia("impdebsol") = NulosN(Fg1.TextMatrix(A, 3))
'''        RstDia("imphabsol") = NulosN(Fg1.TextMatrix(A, 4))
'''
'''        RstDia("impdebdol") = NulosN(Fg1.TextMatrix(A, 9))
'''        RstDia("imphabdol") = NulosN(Fg1.TextMatrix(A, 10))
'''
'''        If mMesActivo = 13 Then
'''            RstDia("fchasi") = CDate("31/12/" + AnoTra)
'''        Else
'''            If mMesActivo = 0 Then
'''                RstDia("fchasi") = CDate("31/12/" + Str(AnoTra - 1))
'''            Else
'''                RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
'''            End If
'''        End If
'''        RstDia("fchdoc") = CDate(TxtFchEmi.Valor)
'''        RstDia("prodiv") = -1
'''        RstDia.Update
'''    Next A
'''
'''
'''
'''    xCon.CommitTrans
'''    Set RstCab = Nothing
'''    Set RstDet = Nothing
'''    Set RstDia = Nothing
'''    MsgBox "La Provición se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito" + vbCr + "Num.Reg. " + Format(mMesActivo, "00") + xNumAsiento, vbInformation, xTitulo
'''
'''    Grabar = True
'''    Me.MousePointer = vbDefault
'''    Exit Function
'''
'''LaCague:
''''    Resume
'''    Me.MousePointer = vbDefault
'''    Grabar = False
'''    xCon.RollbackTrans
'''    Set RstCab = Nothing
'''    Set RstDet = Nothing
'''    Set RstDia = Nothing
'''    MsgBox "No se pudo guardar la provicion por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'''End Function
'''



Private Sub pRecalcularImporte()
    Dim A As Long
    Agregando = True

    For A = 1 To Fg1.Rows - 1
        
        If NulosN(lbl_cb_cod(0).Caption) = 1 Then '--soles
            '--en dolares
            If NulosN(NulosN(TxtTC.Text)) <> 0 Then
                Fg1.TextMatrix(A, 9) = Format(NulosN(Fg1.TextMatrix(A, 3)) / NulosN(TxtTC.Text), FORMAT_MONTO)
                Fg1.TextMatrix(A, 10) = Format(NulosN(Fg1.TextMatrix(A, 4)) / NulosN(TxtTC.Text), FORMAT_MONTO)
            Else
                Fg1.TextMatrix(A, 9) = "0.00"
                Fg1.TextMatrix(A, 10) = "0.00"
            End If
        Else
            '--en soles
            Fg1.TextMatrix(A, 3) = Format(NulosN(Fg1.TextMatrix(A, 9)) * NulosN(TxtTC.Text), FORMAT_MONTO)
            Fg1.TextMatrix(A, 4) = Format(NulosN(Fg1.TextMatrix(A, 10)) * NulosN(TxtTC.Text), FORMAT_MONTO)
        End If
        Agregando = True
        
    Next A
    
    TxtTotDeb.Text = Format(SumaColumna(Fg1, 3), FORMAT_MONTO)
    TxtTotHab.Text = Format(SumaColumna(Fg1, 4), FORMAT_MONTO)
    TxtTotDebDol.Text = Format(SumaColumna(Fg1, 9), FORMAT_MONTO)
    TxtTotHabDol.Text = Format(SumaColumna(Fg1, 10), FORMAT_MONTO)
    
    Agregando = False
    Exit Sub
    
End Sub

Private Sub OpcionesPeriodo()
     
     lblperiodo(0).Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
     lblperiodo(1).Caption = lblperiodo(0).Caption
    '------------------------------------------------------------------------------------------
    '--bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    '------------------------------------------------------------------------------------------
    TDB_FiltroLimpiar Dg1
    Set RstFrm = Nothing
    '------------------------------------------
    
    On Error GoTo error
    Dim nSQL  As String

    nSQL = "SELECT con_proviciones.*, mae_libros.descripcion AS desclib, mae_documento.abrev AS destipdoc, con_meses.descripcion AS descmes, " _
        + vbCr + " (SELECT Sum([con_provicionesdet]![imp]) AS totdeb  FROM con_provicionesdet WHERE (((con_provicionesdet.id)=con_proviciones.id)  AND ((con_provicionesdet.tipo)=0))) & '' AS totdeb1, " _
        + vbCr + " (SELECT Sum([con_provicionesdet]![imp]) AS totdeb FROM con_provicionesdet WHERE  (((con_provicionesdet.id)=con_proviciones.id)   AND ((con_provicionesdet.tipo)=-1))) & '' AS tothab1, " _
        + vbCr + " mae_moneda.descripcion AS mondesc, mae_moneda.simbolo, [con_proviciones]![numser]+'-'+[con_proviciones]![numdoc] AS numedoc, mae_librossub.descripcion AS sublibdesc, " _
        + vbCr + " Format([con_proviciones].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & Mid([con_proviciones].[numreg],3) AS registro " _
        + vbCr + " FROM ((((con_proviciones LEFT JOIN mae_libros ON con_proviciones.idlib = mae_libros.id) LEFT JOIN con_meses ON con_proviciones.idmes = con_meses.id) LEFT JOIN mae_moneda ON con_proviciones.idmon = mae_moneda.id) LEFT JOIN mae_documento ON con_proviciones.tipdoc = mae_documento.id) LEFT JOIN mae_librossub ON con_proviciones.idsublib = mae_librossub.id " _
        + vbCr + " Where (((con_proviciones.ano) = " & AnoTra & "  ) And ((con_proviciones.idmes) = " & mMesActivo & "  )) " _
        + vbCr + " ORDER BY con_proviciones.numreg desc;"
    
    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQL, xCon
    
    Set Dg1.DataSource = RstFrm
    Me.MousePointer = vbDefault
Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pCargarGrid"
    
End Sub






Private Sub CmdImportar_Click()
    '===================================================================================================
    'Creado : 24/02/10 Por: Johan Castro
    'Propósito: Importar un asiento desde excel
    '
    'Entradas:  El usuario seleccionar el archivo
    '
    'Resultados: Asiento en seven expresado a dos monedas
    '
    'Nota:       1.- Seleccionar la moneda
    '            2.- Indicar la fecha
    '            3.- Importar
    '            4.- Formato de excel:
    '                Fila=1;
    '                Col1=Cuenta; Col2=Descripción Cuenta; Col3=Imp Debe; Col4=Imp Haber
    '===================================================================================================
    If QueHace = 3 Then Exit Sub

    '--verificar si selecciono la moneda
    If NulosN(txt_cb(0).Text) = 0 Then
        MsgBox "Falta especificar la moneda", vbInformation, xTitulo
        txt_cb(0).SetFocus
        Exit Sub
    End If
    
    '--verificar si selecciona la fecha, util para expresar a dos monedas
    If IsDate(TxtFchEmi.Valor) = False Then
        MsgBox "Falta especificar la fecha de emisión", vbInformation, xTitulo
        TxtFchEmi.SetFocus
        Exit Sub
    End If
    
    '-----------------------------------------------------------------------------------------------------
    '--muestra ventana para seleccionar archivo
    Dim RutaArchivo As String
    CommonDialog1.FileName = ""
    CommonDialog1.DefaultExt = "*.xls"
    'CommonDialog1.Filter = "Cardfile (*.crd)|*.crd|Textos (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
    CommonDialog1.Filter = "Documentos de Excel (*.xls)|*.xls"
    CommonDialog1.ShowOpen
    If Err Then
        'Cancelada la operación de abrir
        Exit Sub
    Else
        RutaArchivo = CommonDialog1.FileName
    End If
    
    If RutaArchivo = "" Then Exit Sub
    '-----------------------------------------------------------------------------------------------------
    
        
    Dim objExcel As Object '--archivo de excel
    Dim RstPlanCta As New ADODB.Recordset '--relacion de plan de cuenta
    Dim A&
    Dim xNumFilas& '--obtener el numero de registros en el asiento
    Dim sImpDeb As Double '--importe de la columna debe
    Dim sImpHab As Double '--importe de la columna haber
    Dim fDuplicados As Boolean 'indica si hay cuentas duplicadas para mostrar mensaje
    
    Me.MousePointer = vbHourglass
    
    On Error GoTo error
    
    Set objExcel = CreateObject("Excel.Application")
    'Dim objExcel As New Excel.Application
    
    objExcel.Visible = True
    objExcel.SheetsInNewWorkbook = 1
    
    'abre el Libro
    objExcel.WindowState = 2
    objExcel.Workbooks.Open RutaArchivo
    
    Frame10.Left = 4290
    Frame10.Top = 2910
    Label7.Caption = "Cargando registros para la importación"
    Frame10.Visible = True
    
    xNumFilas = 1
    
    '--limpiar grilla para insertar registros
    Fg1.Rows = 1
    DoEvents
    
    '--obtener relacion del plan de cuenta
    RST_Busq RstPlanCta, "select * from con_planctas", xCon
    
    
    With objExcel.ActiveSheet
        
        'DETERMINAMOS EL NUMERO DE FILAS CON DATOS
        A = 2
        Do While NulosC(.Cells(A, 1)) <> ""
            If NulosC(.Cells(A, 1)) <> "" Then
                xNumFilas = xNumFilas + 1
            Else
                Exit Do
            End If
            A = A + 1
        Loop
        
        
        xNumFilas = xNumFilas + 1
        ProgressBar2.Max = xNumFilas
        A = 2
        '--indicar que se esta agregando registros
        Agregando = True
        
        Do While NulosC(.Cells(A, 1)) <> ""
        
            ProgressBar2.Value = A
            DoEvents
            
            Fg1.Rows = Fg1.Rows + 1
                
            RstPlanCta.Filter = ""
            '--filtrar cuenta
            RstPlanCta.Filter = "cuenta='" & NulosC(.Cells(A, 1)) & "" & "'"
            '--verificar si esiste la cuenta
            If RstPlanCta.RecordCount <> 0 Then
                '--colocar datos de la cuenta
                Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(RstPlanCta("cuenta"))
                Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstPlanCta("descripcion"))
                Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosN(RstPlanCta("id"))
                
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosN(RstPlanCta("documentar"))
                Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosN(RstPlanCta("idmodulo"))
            End If
            '--obteniendo los importes
            sImpDeb = NulosN(.Cells(A, 3))
            sImpHab = NulosN(.Cells(A, 4))
            
            If NulosN(txt_cb(0).Text) = 1 Then
                '--importes en MN
                Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(sImpDeb, FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(sImpHab, FORMAT_MONTO)
                
                If ChkAjusteDifCambio.Value = 0 Then
                    '--expresar importes en ME
                    If NulosN(TxtTC.Text) <> 0 Then
                        Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(sImpDeb / NulosN(TxtTC.Text), FORMAT_MONTO)
                        Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(sImpHab / NulosN(TxtTC.Text), FORMAT_MONTO)
                    End If
                End If
            Else
                '--importes en ME
                Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(sImpDeb, FORMAT_MONTO)
                Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(sImpHab, FORMAT_MONTO)
                                
                If ChkAjusteDifCambio.Value = 0 Then
                    '--expresar importes en MN
                    Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(sImpDeb * NulosN(TxtTC.Text), FORMAT_MONTO)
                    Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(sImpHab * NulosN(TxtTC.Text), FORMAT_MONTO)
                End If
            End If
                
            DoEvents
            A = A + 1
        Loop
    End With
        
    '--totalizando importes por columna
    HallarTotal
    
'''    Label7.Caption = "Revizando Cuentas Duplicadas"
    DoEvents
'''    '--validar duplicados de letras
    If Fg1.Rows > Fg1.FixedRows Then
        ProgressBar2.Max = Fg1.Rows - 1
        For A = 1 To Fg1.Rows - 1
            ProgressBar2.Value = A
            DoEvents
            If Fg1.FindRow(Fg1.TextMatrix(A, 1), A + 1, 1, False, False) <> -1 Then
                '--colocar como duplicado
                FORMATO_CELDA Fg1, A, 1, vbBlack, True
                FORMATO_CELDA Fg1, A, 2, vbBlack, True
                FORMATO_CELDA Fg1, A, 3, vbBlack, True
                FORMATO_CELDA Fg1, A, 4, vbBlack, True
                FORMATO_CELDA Fg1, A, 9, vbBlack, True
                FORMATO_CELDA Fg1, A, 10, vbBlack, True

                fDuplicados = True
            End If
        Next A
    End If
    
    Agregando = False
    Frame10.Visible = False
    
    
    objExcel.WindowState = 2
    objExcel.Workbooks.Close
    
    Set objExcel = Nothing
    Set RstPlanCta = Nothing
    
    MsgBox "El proceso terminó de cargar los datos con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    If fDuplicados = True Then
        MsgBox "Hay Cuentas repetidas, estas se muestran en blanco", vbInformation, xTitulo
    End If
    
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Frame10.Visible = False
    
    Me.MousePointer = vbDefault
    Set RstPlanCta = Nothing
    Set objExcel = Nothing
    SHOW_ERROR Me.Name, "CargaDocumentos"
End Sub
