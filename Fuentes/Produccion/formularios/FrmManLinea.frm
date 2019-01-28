VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManLinea 
   Caption         =   "Producción - Configurar Linea de Produccion"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
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
      Height          =   2820
      Left            =   2250
      TabIndex        =   20
      Top             =   3450
      Visible         =   0   'False
      Width           =   6990
      Begin VB.Frame Frame3 
         Height          =   2355
         Left            =   5430
         TabIndex        =   30
         Top             =   320
         Width           =   1485
         Begin VB.CommandButton Cmd 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   330
            Index           =   7
            Left            =   90
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "Cancelar Seleccion"
            Top             =   1920
            Width           =   1300
         End
         Begin VB.CommandButton Cmd 
            Caption         =   "Aceptar"
            Enabled         =   0   'False
            Height          =   330
            Index           =   6
            Left            =   90
            TabIndex        =   34
            ToolTipText     =   "Aceptar Seleccion"
            Top             =   1530
            Width           =   1300
         End
         Begin VB.CommandButton Cmd 
            Caption         =   "&Bajar"
            Enabled         =   0   'False
            Height          =   330
            Index           =   10
            Left            =   90
            TabIndex        =   33
            ToolTipText     =   "Posiciona mas abajo la Tarea"
            Top             =   910
            Width           =   1300
         End
         Begin VB.CommandButton Cmd 
            Caption         =   "&Subir"
            Enabled         =   0   'False
            Height          =   330
            Index           =   9
            Left            =   90
            TabIndex        =   32
            ToolTipText     =   "Posiciona mas arriba la Tarea"
            Top             =   500
            Width           =   1300
         End
         Begin VB.CommandButton Cmd 
            Caption         =   "&Seleccionar"
            Enabled         =   0   'False
            Height          =   330
            Index           =   8
            Left            =   90
            TabIndex        =   31
            ToolTipText     =   "Selecciona la Tarea"
            Top             =   120
            Width           =   1300
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg 
         Height          =   2220
         Index           =   2
         Left            =   90
         TabIndex        =   21
         Top             =   405
         Width           =   5280
         _cx             =   9313
         _cy             =   3916
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
         Rows            =   5
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmManLinea.frx":0000
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
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   6690
         Picture         =   "FrmManLinea.frx":00B2
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   22
         ToolTipText     =   "Cerrar"
         Top             =   60
         Width           =   195
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   8295
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   6960
         Y1              =   2790
         Y2              =   2790
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccion de Tareas"
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
         Left            =   50
         TabIndex        =   23
         Top             =   70
         Width           =   1770
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   0
         X1              =   6960
         X2              =   6960
         Y1              =   0
         Y2              =   2790
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   15
         Top             =   45
         Width           =   6915
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
            Picture         =   "FrmManLinea.frx":039E
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLinea.frx":08E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLinea.frx":0C74
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLinea.frx":0DF8
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLinea.frx":124C
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLinea.frx":1364
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLinea.frx":18A8
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLinea.frx":1DEC
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLinea.frx":1F00
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLinea.frx":2014
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLinea.frx":2468
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLinea.frx":25D4
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLinea.frx":2B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLinea.frx":2EAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManLinea.frx":31C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6915
      Left            =   0
      TabIndex        =   1
      Top             =   375
      Width           =   11175
      _cx             =   19711
      _cy             =   12197
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
      CurrTab         =   1
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
         Height          =   6495
         Left            =   45
         TabIndex        =   5
         Top             =   375
         Width           =   11085
         Begin VB.Frame FrmLinea 
            Caption         =   "[ Linea ]"
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
            Height          =   1635
            Left            =   60
            TabIndex        =   17
            Top             =   1290
            Width           =   10920
            Begin VB.Frame FrmLineaBot 
               Height          =   1425
               Left            =   9090
               TabIndex        =   24
               Top             =   150
               Width           =   1755
               Begin VB.CommandButton Cmd 
                  Caption         =   "&Activar"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   1
                  Left            =   150
                  TabIndex        =   27
                  ToolTipText     =   "Establece como Principal la Linea Actual "
                  Top             =   120
                  Width           =   1400
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "&Agregar"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   2
                  Left            =   150
                  TabIndex        =   26
                  ToolTipText     =   "Carga los valores correspondientes en la Receta"
                  Top             =   540
                  Width           =   1400
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "&Eliminar"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   3
                  Left            =   150
                  TabIndex        =   25
                  ToolTipText     =   "Procesa los valores de Linea"
                  Top             =   990
                  Width           =   1400
               End
            End
            Begin VSFlex7Ctl.VSFlexGrid Fg 
               Height          =   1300
               Index           =   0
               Left            =   120
               TabIndex        =   18
               Top             =   250
               Width           =   8895
               _cx             =   15690
               _cy             =   2293
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
               Rows            =   4
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManLinea.frx":34E2
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
         Begin VB.Frame FrmDetalle 
            Caption         =   "[ Detalle ]"
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
            Height          =   3375
            Left            =   60
            TabIndex        =   13
            Top             =   3030
            Width           =   10920
            Begin VB.OptionButton OptVista 
               Caption         =   "Vista Detallada"
               Height          =   195
               Index           =   1
               Left            =   1740
               TabIndex        =   29
               Top             =   300
               Width           =   1425
            End
            Begin VB.OptionButton OptVista 
               Caption         =   "Vista Resumida"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   28
               Top             =   300
               Width           =   1485
            End
            Begin VB.CommandButton Cmd 
               Caption         =   "&Simular"
               Enabled         =   0   'False
               Height          =   330
               Index           =   4
               Left            =   7830
               TabIndex        =   16
               ToolTipText     =   "Procesa los valores de Linea"
               Top             =   180
               Width           =   1400
            End
            Begin VB.CommandButton Cmd 
               Caption         =   "&Cargar en Receta"
               Enabled         =   0   'False
               Height          =   330
               Index           =   5
               Left            =   9300
               TabIndex        =   15
               ToolTipText     =   "Carga los valores correspondientes en la Receta"
               Top             =   180
               Width           =   1545
            End
            Begin VSFlex7Ctl.VSFlexGrid Fg 
               Height          =   2670
               Index           =   1
               Left            =   60
               TabIndex        =   19
               Top             =   600
               Width           =   10770
               _cx             =   18997
               _cy             =   4710
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
               Rows            =   5
               Cols            =   20
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManLinea.frx":360E
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
         Begin VB.Frame FrmReceta 
            Caption         =   "[ Receta ]"
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
            Height          =   735
            Left            =   60
            TabIndex        =   8
            Top             =   450
            Width           =   10920
            Begin VB.CommandButton Cmd 
               Enabled         =   0   'False
               Height          =   240
               Index           =   0
               Left            =   1980
               Picture         =   "FrmManLinea.frx":3867
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   345
               Width           =   225
            End
            Begin VB.TextBox TxtCodRec 
               Height          =   300
               Left            =   1020
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   0
               Text            =   "TxtCodRec"
               Top             =   315
               Width           =   1215
            End
            Begin VB.Label LblIdRec 
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "LblIdRec"
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
               Height          =   300
               Left            =   9600
               TabIndex        =   11
               Top             =   330
               Width           =   1185
            End
            Begin VB.Label lbl_cb_capt 
               AutoSize        =   -1  'True
               Caption         =   "Descripción"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   10
               Top             =   360
               Width           =   840
            End
            Begin VB.Label LblDetalle 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblDetalle"
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
               Left            =   2220
               TabIndex        =   12
               Top             =   315
               Width           =   8580
            End
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
            TabIndex        =   14
            Top             =   3435
            Width           =   6330
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Linea"
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
            TabIndex        =   6
            Top             =   100
            Width           =   10980
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6495
         Left            =   -11730
         TabIndex        =   2
         Top             =   375
         Width           =   11085
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6000
            Left            =   30
            TabIndex        =   3
            Top             =   495
            Width           =   11010
            _ExtentX        =   19420
            _ExtentY        =   10583
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Id"
            Columns(0).DataField=   "idlinea"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Descripción"
            Columns(1).DataField=   "descripcion"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Receta"
            Columns(2).DataField=   "codrec"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=10451"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=10372"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=3228"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=3149"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=58,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Named:id=33:Normal"
            _StyleDefs(49)  =   ":id=33,.parent=0"
            _StyleDefs(50)  =   "Named:id=34:Heading"
            _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(52)  =   ":id=34,.wraptext=-1"
            _StyleDefs(53)  =   "Named:id=35:Footing"
            _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(55)  =   "Named:id=36:Selected"
            _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=37:Caption"
            _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(59)  =   "Named:id=38:HighlightRow"
            _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(61)  =   "Named:id=39:EvenRow"
            _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(63)  =   "Named:id=40:OddRow"
            _StyleDefs(64)  =   ":id=40,.parent=33"
            _StyleDefs(65)  =   "Named:id=41:RecordSelector"
            _StyleDefs(66)  =   ":id=41,.parent=34"
            _StyleDefs(67)  =   "Named:id=42:FilterBar"
            _StyleDefs(68)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Linea"
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
            TabIndex        =   4
            Top             =   100
            Width           =   11010
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11205
      _ExtentX        =   19764
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
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Actualizar Costo en Lote"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Equivalencia de Costo en Horas"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
Attribute VB_Name = "FrmManLinea"
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
Dim RstValores As New ADODB.Recordset
Dim mRowAdd As Double                   ' identificador unico por fila cuando se agrege una unidad
Dim IdMenuActivo As Integer             'INDICA EL CODIGO DEL MENU ACTIVO
Dim xHorIni As Date                     'ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim cSQL As String
Dim IDLINEA_ As Double
'Para mover el frame
Dim OrigFX As Long
Dim OrigFY As Long

'*****************************************************************************************************
'* Nombre           : Filtrar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA UN FILTRO SOBRE EL RECORDSET RstFrm
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Filtrar()
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(2, 4) As String
   
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":        xCampos(0, 2) = "C":         xCampos(0, 3) = "4200"
    xCampos(1, 0) = "Codigo":       xCampos(1, 1) = "codrec":             xCampos(1, 2) = "C":         xCampos(1, 3) = "1500"
    
    TabOne1.CurrTab = 0
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg1
End Sub

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

Private Sub cmd_Click(Index As Integer)
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    Dim SELECCIONADO As Double
    Dim DESCRIPCION As String
    Dim IDREC As Double
    Dim IDTAR As Double
    Dim IDUNIMED As Double
    Dim FILA As Integer
    Dim Rpta As Integer
    
    Dim SELECCIONADO_AUX As Double
    Dim DESCRIPCION_AUX As String
    Dim IDREC_AUX As Double
    Dim IDTAR_AUX As Double
    Dim IDUNIMED_AUX As Double
    Dim nSQLId As String
    
    If QueHace = 3 Then Exit Sub
    
    Select Case Index
        Case 0 ' Elegir Receta
            ReDim xCampos(3, 4) As String
            Dim nTitulo As String
            Dim xRsAux As New ADODB.Recordset
            
            Set xRs = Nothing
            
            nTitulo = "Recetas"
            
            ' generar la lista de recetas para no considerar en la lista
            cSQL = "SELECT pro_linea.idrec " _
                + vbCr + "From pro_linea " _
                + vbCr + "GROUP BY pro_linea.idrec;"
            
            RST_Busq xRsAux, cSQL, xCon
            ' Se verifica que no se agregue una receta ya existente
            nSQLId = GENERAR_SQL_ID_RST(xRsAux, "idrec", " AND pro_receta.id", "NOT IN", True)

            cSQL = "SELECT pro_receta.id, pro_receta.codrec, alm_inventario.descripcion, mae_familia.descripcion AS desfam, IIf([prirec]=1,'PRINCIPAL','AUXILIAR') AS prioridad " _
                + vbCr + "FROM pro_receta LEFT JOIN (alm_inventario LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id) ON pro_receta.iditem = alm_inventario.id " _
                + vbCr + "Where (((alm_inventario.tippro) = 3)) " & nSQLId _
                + vbCr + "GROUP BY pro_receta.id, pro_receta.codrec, alm_inventario.descripcion, mae_familia.descripcion, IIf([prirec]=1,'PRINCIPAL','AUXILIAR');"
            
            RST_Busq xRs, cSQL, xCon
            
            'descripcion                        'campo                           'tamaño                    'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":    xCampos(0, 3) = "C"
            xCampos(1, 0) = "Cod. Rec":         xCampos(1, 1) = "codrec":        xCampos(1, 2) = "1000":    xCampos(1, 3) = "C"
            xCampos(2, 0) = "Prioridad":        xCampos(2, 1) = "prioridad":     xCampos(2, 2) = "2000":    xCampos(2, 3) = "C"

            CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), "Buscando " & nTitulo, "descripcion", "descripcion", Principio

            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            If NulosC(TxtCodRec.Text) <> "" And RstValores.RecordCount <> 0 Then
                If MsgBox("Esta seguro que desea cambiar de Receta, se eliminara la Linea relacionada a la Receta anterior?" _
                                    , vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
            End If
            
            limpiarRST RstValores, True
            fg(0).Rows = fg(0).FixedRows
            fg(1).Rows = fg(1).FixedRows
                        
            TxtCodRec.Text = NulosC(xRs("codrec")) ' CODIGO DE LA RECETA
            lblIdRec.Caption = NulosC(xRs("id")) ' ID DE LA RECETA
            LblDetalle = NulosC(xRs("descripcion")) ' DESCRIPCION DE LA RECETA

        Case 1 ' Activar Linea
            Dim fila_seleccionada As Integer
            
            fila_seleccionada = fg(0).Row
            ' Se deseleccionan todas las filas
            For A = 1 To fg(0).Rows - 1
                fg(0).TextMatrix(A, 7) = 0
            Next A
            ' se selecciona lo escogido
            fg(0).TextMatrix(fila_seleccionada, 7) = -1
            
        Case 2 ' Agregar Linea
            Set xRs = Nothing
            
            If NulosN(lblIdRec.Caption) = 0 Then Exit Sub
        
            cSQL = "SELECT pro_recetatar.idrec, pro_recetatar.orden, pro_tareas.id, pro_tareas.descripcion, pro_tareas.idunimed " _
                + vbCr + "FROM pro_recetatar LEFT JOIN pro_tareas ON pro_recetatar.idtar = pro_tareas.id " _
                + vbCr + "Where (((pro_recetatar.idrec) = " & NulosN(lblIdRec.Caption) & ")) " _
                + vbCr + "GROUP BY pro_recetatar.idrec, pro_recetatar.orden, pro_tareas.id, pro_tareas.descripcion, pro_tareas.idunimed " _
                + vbCr + "ORDER BY pro_recetatar.orden;"
                
            RST_Busq xRs, cSQL, xCon
            
            fg(2).Rows = 1
            If xRs.State = 0 Then GoTo ERROR_AL_ENCONTRAR_TAREA
            If xRs.RecordCount = 0 Then GoTo ERROR_AL_ENCONTRAR_TAREA
            
            xRs.MoveFirst
            For A = 1 To xRs.RecordCount
                fg(2).Rows = fg(2).Rows + 1
                
                fg(2).TextMatrix(A, 1) = -1
                fg(2).TextMatrix(A, 2) = xRs("descripcion")
                fg(2).TextMatrix(A, 3) = xRs("idrec")
                fg(2).TextMatrix(A, 4) = xRs("id")
                fg(2).TextMatrix(A, 5) = xRs("idunimed")
                
                xRs.MoveNext
                If xRs.EOF Then Exit For
            Next A

            Frame4.Visible = True
            fg(2).SetFocus
            fg(2).Select 1, 1
            Exit Sub


ERROR_AL_ENCONTRAR_TAREA:
            MsgBox "La Receta procesada no tiene Tareas, procese las Tareas en Receta para agregarlas", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Set xRs = Nothing
            
        Case 3 ' Eliminar Linea
            Dim IDLINEA_AUX As Double
            
            Rpta = MsgBox("Esta seguro de eliminar la Linea seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                IDLINEA_AUX = fg(0).TextMatrix(fg(0).Row, 8)
                If IDLINEA_AUX = IDLINEA_ Then IDLINEA_ = IDLINEA_ - 1
                
                RstValores.Filter = "idlineadet = " & IDLINEA_AUX
                
                limpiarRST RstValores, False
                fg(0).RemoveItem fg(0).Row
                
                fg_RowColChange 0
            End If
            
        Case 4 ' Simular Linea
            procesarLinea False, True, False, False
            
        Case 5 ' Cargar en Receta
        
        Case 6 ' Aceptar Seleccion de Tareas
            Dim contador As Integer
            
            If NulosN(IDLINEA_) = 0 Then
                IDLINEA_ = HallaCodigoTabla("pro_linea", xCon, "id")
            Else
                IDLINEA_ = IDLINEA_ + 1
            End If
            
            If fg(2).Rows <= 1 Then Exit Sub
            contador = 0
            fg(1).Rows = 1
            RstValores.Filter = adFilterNone
            
            For A = 1 To fg(2).Rows - 1
                ' Si esta seleccionado
                If fg(2).TextMatrix(A, 1) = -1 Then
                    contador = contador + 1
                    RstValores.AddNew
                    
                    RstValores("orden") = contador
                    RstValores("descripcion") = fg(2).TextMatrix(A, 2)
                    RstValores("idlineadet") = IDLINEA_
                    RstValores("idtar") = fg(2).TextMatrix(A, 4)
                    RstValores("idunimed") = fg(2).TextMatrix(A, 5)
                    RstValores("prioridad") = 1
                    
                    RstValores.Update
                End If
            Next A
            
            fg(0).Rows = fg(0).Rows + 1
            fg(0).TextMatrix(fg(0).Rows - 1, 8) = IDLINEA_
            
            Dim RstAux As New ADODB.Recordset
            Set RstAux = Nothing

            cSQL = "SELECT pro_receta.id AS idrec, mae_unidades.abrev, mae_unidades.id As idunimed " _
                + vbCr + "FROM (pro_receta LEFT JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id " _
                + vbCr + "Where (((pro_receta.id) = " & NulosN(lblIdRec.Caption) & ")) " _
                + vbCr + "GROUP BY pro_receta.id, mae_unidades.abrev, mae_unidades.id;"
            
            RST_Busq RstAux, cSQL, xCon
            
            If RstAux.State = 0 Then Exit Sub
            If RstAux.RecordCount = 0 Then fg(0).TextMatrix(fg(0).Rows - 1, 2) = "Kg."
            
            fg(0).TextMatrix(fg(0).Rows - 1, 2) = RstAux("abrev")
            fg(0).TextMatrix(fg(0).Rows - 1, 9) = RstAux("idunimed")
            
            Frame4.Visible = False

            fg(0).SetFocus
            fg(0).Select fg(0).Rows - 1, 1
            
        Case 7 ' Cancelar Seleccion de Tareas
            fg(2).Rows = 1
            Frame4.Visible = False
            
        Case 8 ' Seleccionar una Tarea
            fg(2).TextMatrix(fg(2).Row, 1) = -1
            
        Case 9 ' Subir una posicion una Tarea
            SELECCIONADO = fg(2).TextMatrix(fg(2).Row, 1)
            DESCRIPCION = fg(2).TextMatrix(fg(2).Row, 2)
            IDREC = fg(2).TextMatrix(fg(2).Row, 3)
            IDTAR = fg(2).TextMatrix(fg(2).Row, 4)
            IDUNIMED = fg(2).TextMatrix(fg(2).Row, 5)
            FILA = fg(2).Row
            
            If FILA = 1 Then Exit Sub
            
            SELECCIONADO_AUX = fg(2).TextMatrix(fg(2).Row - 1, 1)
            DESCRIPCION_AUX = fg(2).TextMatrix(fg(2).Row - 1, 2)
            IDREC_AUX = fg(2).TextMatrix(fg(2).Row - 1, 3)
            IDTAR_AUX = fg(2).TextMatrix(fg(2).Row - 1, 4)
            IDUNIMED_AUX = fg(2).TextMatrix(fg(2).Row - 1, 5)
            
            fg(2).TextMatrix(FILA - 1, 1) = SELECCIONADO
            fg(2).TextMatrix(FILA - 1, 2) = DESCRIPCION
            fg(2).TextMatrix(FILA - 1, 3) = IDREC
            fg(2).TextMatrix(FILA - 1, 4) = IDTAR
            fg(2).TextMatrix(FILA - 1, 5) = IDUNIMED
            
            fg(2).TextMatrix(FILA, 1) = SELECCIONADO_AUX
            fg(2).TextMatrix(FILA, 2) = DESCRIPCION_AUX
            fg(2).TextMatrix(FILA, 3) = IDREC_AUX
            fg(2).TextMatrix(FILA, 4) = IDTAR_AUX
            fg(2).TextMatrix(FILA, 5) = IDUNIMED_AUX
            
            fg(2).Select FILA - 1, 2
            
        Case 10 ' Bajar una posicion una Tarea
            SELECCIONADO = fg(2).TextMatrix(fg(2).Row, 1)
            DESCRIPCION = fg(2).TextMatrix(fg(2).Row, 2)
            IDREC = fg(2).TextMatrix(fg(2).Row, 3)
            IDTAR = fg(2).TextMatrix(fg(2).Row, 4)
            IDUNIMED = fg(2).TextMatrix(fg(2).Row, 5)
            FILA = fg(2).Row
            
            If FILA = fg(2).Rows - 1 Then Exit Sub
            
            SELECCIONADO_AUX = fg(2).TextMatrix(fg(2).Row + 1, 1)
            DESCRIPCION_AUX = fg(2).TextMatrix(fg(2).Row + 1, 2)
            IDREC_AUX = fg(2).TextMatrix(fg(2).Row + 1, 3)
            IDTAR_AUX = fg(2).TextMatrix(fg(2).Row + 1, 4)
            IDUNIMED_AUX = fg(2).TextMatrix(fg(2).Row + 1, 5)
            
            fg(2).TextMatrix(FILA + 1, 1) = SELECCIONADO
            fg(2).TextMatrix(FILA + 1, 2) = DESCRIPCION
            fg(2).TextMatrix(FILA + 1, 3) = IDREC
            fg(2).TextMatrix(FILA + 1, 4) = IDTAR
            fg(2).TextMatrix(FILA + 1, 5) = IDUNIMED
            
            fg(2).TextMatrix(FILA, 1) = SELECCIONADO_AUX
            fg(2).TextMatrix(FILA, 2) = DESCRIPCION_AUX
            fg(2).TextMatrix(FILA, 3) = IDREC_AUX
            fg(2).TextMatrix(FILA, 4) = IDTAR_AUX
            fg(2).TextMatrix(FILA, 5) = IDUNIMED_AUX
            
            fg(2).Select FILA + 1, 2
    End Select
End Sub

Private Sub procesarLinea(Optional MOSTRAR_ As Boolean = True, Optional PROCESAR_ As Boolean = False, _
                            Optional CARGAR_ As Boolean = False, Optional GRABAR_ As Boolean = False)
    
    Dim A As Integer
    Dim RstLinea As New ADODB.Recordset
    Dim cSQL As String
    
    If MOSTRAR_ Then ' Mostrar la linea activa
    End If
    
    If PROCESAR_ Then ' Procesar linea
        Dim cantMP As Double
        Dim error As Boolean
        
On Error GoTo ERROR_AL_PROCESAR
        cantMP = NulosN(fg(0).TextMatrix(fg(0).Row, 3))
        error = False
        
        With fg(1)
            For A = 1 To .Rows - 1
                If NulosN(cantMP) = 0 Then error = True
                If .TextMatrix(A, 3) = "" Then error = True
                If .TextMatrix(A, 5) = "" Then error = True
                If .TextMatrix(A, 7) = "" Then error = True
            Next A
            
            If error Then
                MsgBox "No ha ingresado correctamente toda la informacion necesaria", vbExclamation, xTitulo
                Exit Sub
            End If
            
            Dim SUMA_TIEMPO As Double ' Indica el tiempo transcurrido
            Dim SUMA_OPERARIOS As Double ' Indica el total de operarios
            Dim SUMA_EFICIENCIA As Double ' Indica el total de operarios
            
            SUMA_TIEMPO = 0
            SUMA_OPERARIOS = 0
            SUMA_EFICIENCIA = 0
            For A = 1 To .Rows - 1
                '*************************************************************Cantidad Procesada
                If A = 1 Then
                    .TextMatrix(A, 4) = Format((NulosN(.TextMatrix(A, 3)) * NulosN(cantMP)) / 100, "0.00")
                Else
                    .TextMatrix(A, 4) = Format((NulosN(.TextMatrix(A, 3)) * NulosN(.TextMatrix(A - 1, 4))) / 100, "0.00")
                End If
                
                '*************************************************************Personal Ideal
                .TextMatrix(A, 6) = Format(NulosN(.TextMatrix(A, 4)) / NulosN(.TextMatrix(A, 5)), "0.00")
                
                '*************************************************************Personal Real en planta
                If .TextMatrix(A, 7) = 1 Then
                    .TextMatrix(A, 8) = Format(NulosN(.TextMatrix(A, 6)), "00")
                Else
                    .TextMatrix(A, 8) = Format(NulosN(.TextMatrix(A, 6)) + 1, "00")
                End If
                If NulosN(.TextMatrix(A, 8)) = 0 Then .TextMatrix(A, 8) = 1
                
                SUMA_OPERARIOS = SUMA_OPERARIOS + NulosN(.TextMatrix(A, 8))
                
                '*************************************************************Eficiencia por personal
                .TextMatrix(A, 9) = Format((NulosN(.TextMatrix(A, 6)) / NulosN(.TextMatrix(A, 8))) * 100, "0.00")
                
                '*************************************************************Eficiencia por tarea
                .TextMatrix(A, 10) = Format((NulosN(.TextMatrix(A, 9)) * NulosN(.TextMatrix(A, 8))), "0.00")
                SUMA_EFICIENCIA = SUMA_EFICIENCIA + NulosN(.TextMatrix(A, 10))
                
                '*************************************************************Avance Real
                .TextMatrix(A, 11) = NulosN(.TextMatrix(A, 5)) * NulosN(.TextMatrix(A, 8))
                
                '*************************************************************Duracion de la tarea
                .TextMatrix(A, 12) = Format(NulosN(.TextMatrix(A, 4)) / NulosN(.TextMatrix(A, 11)), "0.00")
                
                '*************************************************************Desvalance entre horas
                If A = 1 Then
                    .TextMatrix(A, 14) = 0
                Else
                    .TextMatrix(A, 14) = NulosN(.TextMatrix(A, 12)) - NulosN(.TextMatrix(A - 1, 12))
                End If
                
                '*************************************************************Intervalo entre Tareas
                If A = 1 Then
                    .TextMatrix(A, 15) = 0
                Else
                    If .TextMatrix(A, 14) >= 0 Then
                        .TextMatrix(A, 15) = 0
                    Else
                        .TextMatrix(A, 15) = Format(Abs(NulosN(.TextMatrix(A, 14))), "0.00")
                    End If
                End If
                SUMA_TIEMPO = SUMA_TIEMPO + .TextMatrix(A, 15)
                
                '*************************************************************Duracion real de la tarea
                .TextMatrix(A, 16) = SUMA_TIEMPO + NulosN(.TextMatrix(A, 12))
                
                '*************************************************************Factor
                .TextMatrix(A, 13) = Format((NulosN(.TextMatrix(A, 16)) * NulosN(.TextMatrix(A, 8))) / ((NulosN(.TextMatrix(A, 4)) * 100) / NulosN(.TextMatrix(A, 3))), "0.000000")
            Next A
            
            fg(0).TextMatrix(fg(0).Row, 4) = Format(NulosN(.TextMatrix(.Rows - 1, 4)) / NulosN(.TextMatrix(.Rows - 1, 16)), "0.00") ' Unid/hora
            fg(0).TextMatrix(fg(0).Row, 5) = SUMA_OPERARIOS                                     ' Num Operarios
            fg(0).TextMatrix(fg(0).Row, 6) = Format(SUMA_EFICIENCIA / SUMA_OPERARIOS, "0.00")   ' Eficiencia de la Linea
            
            ' Se filtra la linea
            RstValores.Filter = "idlineadet = " & fg(0).TextMatrix(fg(0).Row, 8)
            limpiarRST RstValores, False
            
            If RstValores.State = 0 Then Exit Sub
            For A = 1 To fg(1).Rows - 1
                RstValores.AddNew
                
                RstValores("orden") = fg(1).TextMatrix(A, 1)
                RstValores("descripcion") = fg(1).TextMatrix(A, 2)
                RstValores("rdmto") = fg(1).TextMatrix(A, 3)
                RstValores("kghora") = fg(1).TextMatrix(A, 5)
                RstValores("prioridad") = fg(1).TextMatrix(A, 7)
                RstValores("numop") = fg(1).TextMatrix(A, 8)
                
                RstValores("procesado") = fg(1).TextMatrix(A, 4)
                RstValores("numopideal") = fg(1).TextMatrix(A, 6)
                RstValores("eficop") = fg(1).TextMatrix(A, 9)
                RstValores("efictar") = fg(1).TextMatrix(A, 10)
                RstValores("cantreal") = fg(1).TextMatrix(A, 11)
                RstValores("durtar") = fg(1).TextMatrix(A, 12)
                RstValores("factor") = fg(1).TextMatrix(A, 13)
                RstValores("desvalance") = fg(1).TextMatrix(A, 14)
                RstValores("intervalo") = fg(1).TextMatrix(A, 15)
                RstValores("durtarreal") = fg(1).TextMatrix(A, 16)
                
                RstValores("idlineadet") = fg(1).TextMatrix(A, 17)
                RstValores("idtar") = fg(1).TextMatrix(A, 18)
                RstValores("idunimed") = fg(1).TextMatrix(A, 19)
                
                RstValores.Update
            Next A
            
            pCargarDatosValores
        End With
        Exit Sub
ERROR_AL_PROCESAR:
        MsgBox "Ha ocurrido un error al procesar los Datos, intente otra vez", vbExclamation, xTitulo
        Resume
        Exit Sub
    End If
    
    If CARGAR_ Then ' CARGAR LOS VALORES ADECUADOS A LA RECETA
'        Dim xTiempo As Double
'        Dim xHorEst As String
'
'On Error GoTo ERROR_AL_CARGAR
'        For A = 1 To Fg3.Rows - 1
'            xTiempo = NulosN(Fg4.TextMatrix(A, 14))
'            xHorEst = Format(Int(xTiempo), "00")
'            xHorEst = xHorEst & ":" & Format(((xTiempo * 60) Mod 60), "00")
'            Fg3.TextMatrix(A, 7) = Format(xHorEst, "HH:mm") ' Arranque
'
'            Fg3.TextMatrix(A, 8) = Fg4.TextMatrix(A, 12) 'Factor
'            Fg3.TextMatrix(A, 9) = Fg4.TextMatrix(A, 7) 'numero de operarios
'            Fg3.TextMatrix(A, 11) = Fg4.TextMatrix(A, 2) ' Rendimiento
'        Next A
'        MsgBox "Se ha cargado correctamente la Linea", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Frm3.Visible = False ' Se oculta la linea
'        Exit Sub
'ERROR_AL_CARGAR:
'        MsgBox "Ha ocurrido un error al cargar los Datos, intente otra vez", vbExclamation, xTitulo
'        Exit Sub
    End If
    
    If GRABAR_ Then ' Grabar los Datos de la linea
'        xCon.BeginTrans
'
'On Error GoTo ERROR_AL_GRABAR
'
'        xCon.Execute "DELETE * FROM  pro_recetalinea WHERE idrec = " & NulosN(Fg4.TextMatrix(1, 16)) & ""
'
'        RST_Busq RstLinea, "SELECT TOP 1 * FROM pro_recetalinea", xCon
'
'        For A = 1 To Fg4.Rows - 1
'            RstLinea.AddNew
'            RstLinea("idrec") = NulosN(Fg4.TextMatrix(A, 16))
'            RstLinea("idtar") = NulosN(Fg4.TextMatrix(A, 17))
'            RstLinea("idunimed") = NulosN(Fg4.TextMatrix(A, 18))
'            RstLinea("cantidad") = NulosN(TxtcantMP.Text)
'            RstLinea("rdmto") = NulosN(Fg4.TextMatrix(A, 2))
'            RstLinea("kghora") = NulosN(Fg4.TextMatrix(A, 4))
'            RstLinea("prioridad") = NulosN(Fg4.TextMatrix(A, 6))
'            RstLinea("frechora") = Val(LblFrecLinea.Caption)
'            RstLinea.Update
'        Next A
'
'        xCon.CommitTrans
'        MsgBox "Se ha Grabado correctamente la Linea", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Exit Sub
'ERROR_AL_GRABAR:
'        xCon.RollbackTrans
'        Set RstLinea = Nothing
'        MsgBox "Ha ocurrido un error al Grabar los Datos, intente otra vez", vbExclamation, xTitulo
'        Exit Sub
    End If
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
        If RstFrm.RecordCount = 0 Then Exit Sub
        VerMovimientos1 IdMenuActivo, NulosN(RstFrm("id")), xCon
    End If
End Sub

Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If QueHace = 3 Then Exit Sub
    If Index = 0 Then Exit Sub

    If Index = 1 Then
        RstValores.Filter = "idlineadet = " & fg(1).TextMatrix(fg(1).Row, 17) & _
                                " And idtar = " & fg(1).TextMatrix(fg(1).Row, 18)
                                
        If RstValores.State = 0 Then Exit Sub
        If RstValores.RecordCount = 0 Then Exit Sub
        
        Select Case Col
            Case 3 ' rendimiento
                RstValores("rdmto") = NulosN(fg(1).TextMatrix(Row, Col))
            Case 5 ' unid/hora
                RstValores("kghora") = NulosN(fg(1).TextMatrix(Row, Col))
            Case 7 ' prioridad
                RstValores("prioridad") = NulosN(fg(1).TextMatrix(Row, Col))
        End Select
        RstValores.Update
    End If
End Sub

Private Sub Fg_EnterCell(Index As Integer)
    If QueHace = 3 Then
        fg(Index).Editable = flexEDNone
        fg(Index).SelectionMode = flexSelectionByRow
        Exit Sub
    End If
    fg(Index).Editable = flexEDKbdMouse
    fg(Index).SelectionMode = flexSelectionFree
End Sub

Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    
    If Index = 0 Then
        Select Case Col
            Case 2, 4, 5, 6
                KeyAscii = 0
            Case 3
                If validar_numero(KeyAscii) = False Then KeyAscii = 0
        End Select
    End If
    
    If Index = 1 Then
        Select Case Col
            Case 1, 2, 4, 6, 8, 9, 10, 11, 12, 13, 14, 15, 16
                KeyAscii = 0
            Case 3, 5, 7
                If validar_numero(KeyAscii) = False Then KeyAscii = 0
        End Select
    End If
End Sub

Private Sub Fg_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo error
    If Index = 0 Then Exit Sub
    If KeyCode = 114 Or KeyCode = vbKeyInsert Then cmd_Click 0      'F3 = Agregar Item
    If KeyCode = 115 Or KeyCode = vbKeyDelete Then cmd_Click 1      'F4 = Eliminar Item
    Exit Sub
    
error:
    SHOW_ERROR Me.Name, "Fg_KeyUp (" & Index & ")"
End Sub

Private Sub fg_RowColChange(Index As Integer)
    If Agregando = True Then Exit Sub
    
    If Index = 1 Then Exit Sub
    
    If fg(0).Row < 1 Then
        fg(1).Rows = 1
        Exit Sub
    End If
    
    If RstValores.State = 0 Then Exit Sub
    
    ' Mostramos los insumos de la receta
    RstValores.Filter = adFilterNone
    RstValores.Filter = "idlineadet = " & NulosN(fg(0).TextMatrix(fg(0).Row, 8))
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
                
        cSQL = "SELECT pro_linea.idlinea, alm_inventario.descripcion, pro_receta.id AS idrec, pro_receta.codrec " _
            + vbCr + "FROM (pro_linea LEFT JOIN pro_receta ON pro_linea.idrec = pro_receta.id) LEFT JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id " _
            + vbCr + "GROUP BY pro_linea.idlinea, alm_inventario.descripcion, pro_receta.id, pro_receta.codrec;"
                    
        RST_Busq RstFrm, cSQL, xCon

        Set Dg1.DataSource = RstFrm
        iniciarCampos
    End If
    Exit Sub
error:
    SHOW_ERROR Me.Name, "Form_Activate"
End Sub

Private Sub iniciarCampos()
    ' ocultando las columnas de codigos
    OCULTAR_COL fg(0), 8, 9
    OCULTAR_COL fg(1), 13, 13
    OCULTAR_COL fg(1), 16, 19
    OCULTAR_COL fg(2), 3, 5
    
    fg(0).ColFormat(3) = "0.00"
    fg(0).ColFormat(4) = "0.00"
    fg(0).ColFormat(5) = "00"
    fg(0).ColFormat(6) = "0.00"
    
    fg(1).ColFormat(3) = "0.00"
    fg(1).ColFormat(4) = "0.00"
    fg(1).ColFormat(5) = "00"
    fg(1).ColFormat(6) = "00"
    fg(1).ColFormat(7) = "00"
    fg(1).ColFormat(8) = "00"
    fg(1).ColFormat(9) = "0.00"
    fg(1).ColFormat(10) = "0.00"
    fg(1).ColFormat(11) = "0.00"
    fg(1).ColFormat(12) = "0.00"
    fg(1).ColFormat(13) = "0.00"
    fg(1).ColFormat(14) = "0.00"
    fg(1).ColFormat(15) = "0.00"
    fg(1).ColFormat(16) = "0.00"
    
    TabOne1.CurrTab = 0
    
    fg(1).FrozenCols = 2
    
    fg(0).SelectionMode = flexSelectionByRow
    fg(1).SelectionMode = flexSelectionByRow
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
    Label5.Caption = "Agregando Linea"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    
    If RstValores.State = 0 Then pCargarDatosRstTemp 0
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
    TxtCodRec.Text = ""
    LblDetalle.Caption = ""
    lblIdRec.Caption = ""
    fg(0).Rows = 1
    fg(1).Rows = 1
    fg(2).Rows = 1
    IDLINEA_ = 0
End Sub

'*****************************************************************************************************
'* Nombre           : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES TEXTBOX Y COMMAND
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Bloquea(band As Boolean)
    TxtCodRec.Locked = band
    habilitar Cmd, Not band
    If Not band Then
        fg(0).Editable = flexEDKbdMouse
        fg(1).Editable = flexEDKbdMouse
        fg(2).Editable = flexEDKbdMouse
    Else
        fg(0).Editable = flexEDNone
        fg(1).Editable = flexEDNone
        fg(2).Editable = flexEDNone
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    QueHace = 3
    SeEjecuto = False
    Agregando = False
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
    Label5.Caption = "Detalle de la Linea"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Dg1.SetFocus
    limpiarRST RstValores
End Sub

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub

    If Me.Width <= 12000 Then Me.Width = 12000
    If Me.Height <= 8100 Then Me.Height = 8100

    ' Se dimensiona el Contenido
    TabOne1.Width = Me.Width - 100
    TabOne1.Height = Me.Height - 700
    
    Label4(0).Width = Me.Width - 100
    Dg1.Width = TabOne1.Width - 150
    Dg1.Height = TabOne1.Height - 1150
        
    ' Se dimensiona el Detalle
    ' DETALLE DE RECETA
    Label5.Width = Me.Width - 100
    
    FrmReceta.Top = TabOne1.Top + 100
    FrmReceta.Width = TabOne1.Width - 200
    
    LblDetalle.Width = FrmReceta.Width - 2300
    
    ' DESCRIPCION DE LINEA
    FrmLinea.Top = TabOne1.Top + 1000
    FrmLinea.Width = TabOne1.Width - 200
    
    FrmLineaBot.Left = FrmLinea.Width - 1800
    
    fg(0).Width = FrmLinea.Width - 2000
    
    ' DETALLE LINEA
    FrmDetalle.Top = TabOne1.Top + 2800
    FrmDetalle.Width = TabOne1.Width - 200
    FrmDetalle.Height = TabOne1.Height - 3800
    
    fg(1).Width = FrmDetalle.Width - 150
    fg(1).Height = FrmDetalle.Height - 700
    
    Cmd(4).Left = FrmDetalle.Width - 3200
    Cmd(5).Left = FrmDetalle.Width - 1700
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    If QueHace <> 3 Then
'        MsgBox "No puede salir del formulario mientras este ingresando o modificando un Costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Cancel = True
'        Exit Sub
'    Else
'        Set RstFrm = Nothing
'        SeEjecuto = False
'    End If
'End Sub

Private Sub OptVista_Click(Index As Integer)
    If Index = 0 Then
        OCULTAR_COL fg(1), 11, 16
    End If
    If Index = 1 Then
        fg(1).ColWidth(11) = 1300
        fg(1).ColWidth(12) = 1300
        fg(1).ColWidth(14) = 1300
        fg(1).ColWidth(15) = 1300
        fg(1).ColWidth(16) = 1300
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 3 Then MuestraSegundoTab
    Else
        limpiarRST RstValores, True
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
            RstFrm.Find "idlinea = " & mIdRegistro & ""
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
    If fValidarDatos() = False Then
        xTitulo = "Error"
        MsgBox "No se pudo realizar la Operacion, ingrese correctamente los Datos o Procese la Linea correctamente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Function
    End If
    
    If MsgBox("¿Seguro que desea " + IIf(QueHace = 1, "grabar", "Modificar") + " la Linea?", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo SALIR
       
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xId As Double
    Dim xCol&, xFil&, xCorr&
    Dim A As Integer
    
    
On Error GoTo LaCague

    xCon.BeginTrans
    Me.MousePointer = vbHourglass
    
    If QueHace = 1 Then
        xId = HallaCodigoTabla("pro_linea", xCon, "id")
        RST_Busq RstCab, "SELECT top 1 * FROM pro_linea ", xCon
    Else
        xId = NulosN(RstFrm("idlinea"))
        
        xCon.Execute "DELETE * FROM pro_lineadet WHERE idlinea = " & xId & ""
        xCon.Execute "DELETE * FROM pro_linea WHERE idlinea = " & xId & ""
      
        RST_Busq RstCab, "SELECT top 1 * FROM pro_linea ", xCon
    End If
    
    mIdRegistro = xId
    
    RST_Busq RstDet, "SELECT top 1 * FROM pro_lineadet", xCon
    
    ' Recorrer cabeceras
    For xFil = 1 To fg(0).Rows - 1
        RstCab.AddNew
        RstCab("id") = NulosN(fg(0).TextMatrix(xFil, 8))
        RstCab("idlinea") = xId
        RstCab("idrec") = NulosN(lblIdRec.Caption)
        RstCab("descripcion") = NulosC(fg(0).TextMatrix(xFil, 1))
        RstCab("idunimed") = NulosN(fg(0).TextMatrix(xFil, 9))
        RstCab("cantidad") = NulosN(fg(0).TextMatrix(xFil, 3))
        RstCab("efic") = NulosN(fg(0).TextMatrix(xFil, 6))
        RstCab("numop") = NulosN(fg(0).TextMatrix(xFil, 5))
        RstCab("kghora") = NulosN(fg(0).TextMatrix(xFil, 4))
        RstCab("activo") = NulosN(fg(0).TextMatrix(xFil, 7))
        
        RstCab.Update
        
        RstValores.Filter = "idlineadet= " & NulosN(fg(0).TextMatrix(xFil, 8))
        If RstValores.RecordCount = 0 Then GoTo LaCague
        
        RstValores.MoveFirst
        While Not RstValores.EOF
            RstDet.AddNew
            RstDet("idlinea") = xId
            RstDet("idlineadet") = NulosN(RstValores("idlineadet"))
            RstDet("idrec") = NulosN(lblIdRec.Caption)
            RstDet("orden") = NulosN(RstValores("orden"))
            RstDet("idtar") = NulosN(RstValores("idtar"))
            RstDet("idunimed") = NulosN(RstValores("idunimed"))
            RstDet("rdmto") = NulosN(RstValores("rdmto"))
            RstDet("prioridad") = NulosN(RstValores("prioridad"))
            RstDet("kghora") = NulosN(RstValores("kghora"))
            RstDet("numop") = NulosN(RstValores("numop"))
            RstDet("procesado") = NulosN(RstValores("procesado"))
            RstDet("numopideal") = NulosN(RstValores("numopideal"))
            RstDet("eficop") = NulosN(RstValores("eficop"))
            RstDet("efictar") = NulosN(RstValores("efictar"))
            RstDet("cantreal") = NulosN(RstValores("cantreal"))
            RstDet("durtar") = NulosN(RstValores("durtar"))
            RstDet("factor") = NulosN(RstValores("factor"))
            RstDet("desvalance") = NulosN(RstValores("desvalance"))
            RstDet("intervalo") = NulosN(RstValores("intervalo"))
            RstDet("durtarreal") = NulosN(RstValores("durtarreal"))
            
            RstDet.Update
            RstValores.MoveNext
        Wend
    Next xFil
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    xCon.CommitTrans
    xTitulo = "Grabar"
    MsgBox "La Linea se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    Grabar = True

SALIR:
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing
    Exit Function

LaCague:
    xCon.RollbackTrans
    'Resume
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
    Label5.Caption = "Modificando Linea"
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    limpiarRST RstValores
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
    
    xId = NulosN(RstFrm.Fields("idlinea"))

    Set RstTmp = Nothing
    Rpta = MsgBox("Esta seguro de eliminar la Linea seleccionada", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM pro_lineadet WHERE idlinea =" & xId & " "
        xCon.Execute "DELETE * FROM pro_linea WHERE idlinea =" & xId & " "
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & xId & " AND idform = " & IdMenuActivo
        
        MsgBox "La linea se eliminó con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
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
    Dim A As Integer
    Dim valor As Boolean
    Dim SELECCIONADO As Double
    
    valor = True
    SELECCIONADO = 0
    
    If NulosC(LblDetalle.Caption) = "" Then valor = False
    
    For A = 1 To fg(0).Rows - 1
        If NulosN(fg(0).TextMatrix(A, 4)) = 0 Then valor = False: Exit For
        If NulosN(fg(0).TextMatrix(A, 7)) = -1 Then SELECCIONADO = SELECCIONADO + 1
    Next A
    
    If SELECCIONADO = 0 Or SELECCIONADO > 1 Then valor = False
    
    fValidarDatos = valor
End Function

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
    
    On Error GoTo error
    
    TxtCodRec.Text = RstFrm("codrec")
    LblDetalle.Caption = RstFrm("descripcion")
    lblIdRec.Caption = RstFrm("idrec")
    
    cSQL = "SELECT pro_linea.id, pro_linea.descripcion, mae_unidades.abrev, pro_linea.idunimed, pro_linea.cantidad, pro_linea.kghora, pro_linea.numop, pro_linea.efic, pro_linea.activo " _
        + vbCr + "FROM pro_linea LEFT JOIN mae_unidades ON pro_linea.idunimed = mae_unidades.id " _
        + vbCr + "Where (((pro_linea.idlinea) = " & RstFrm("idlinea") & ")) " _
        + vbCr + "GROUP BY pro_linea.id, pro_linea.descripcion, mae_unidades.abrev, pro_linea.idunimed, pro_linea.cantidad, pro_linea.kghora, pro_linea.numop, pro_linea.efic, pro_linea.activo;"

    RST_Busq RstTmp, cSQL, xCon
    
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
                .TextMatrix(.Rows - 1, 1) = NulosC(RstTmp.Fields("descripcion"))
                .TextMatrix(.Rows - 1, 2) = NulosC(RstTmp.Fields("abrev"))
                .TextMatrix(.Rows - 1, 3) = NulosN(RstTmp.Fields("cantidad"))
                .TextMatrix(.Rows - 1, 4) = NulosN(RstTmp.Fields("kghora"))
                .TextMatrix(.Rows - 1, 5) = NulosN(RstTmp.Fields("numop"))
                .TextMatrix(.Rows - 1, 6) = NulosN(RstTmp.Fields("efic"))
                .TextMatrix(.Rows - 1, 7) = NulosN(RstTmp.Fields("activo"))
                .TextMatrix(.Rows - 1, 8) = NulosN(RstTmp.Fields("id"))
                .TextMatrix(.Rows - 1, 9) = NulosN(RstTmp.Fields("idunimed"))
                
                
                ' cargar datos de los valores de las tareas
                pCargarDatosRstTemp NulosN(fg(0).TextMatrix(fg(0).Rows - 1, 8))
    
                RstTmp.MoveNext
            Loop
        End With
    End If
    
    Set RstTmp = Nothing
    
    
    If fg(0).Rows > 1 Then
        fg(0).Row = 1:  fg(0).Col = 1:
        Agregando = False
        fg_RowColChange 0
        OptVista(0).Value = True
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
    Set RstTmp = Nothing
    
    ' definir la estructura de recordset
    
'SELECT pro_lineadet.idlineadet, pro_lineadet.orden, pro_lineadet.idrec, pro_lineadet.idtar, pro_lineadet.idunimed, pro_lineadet.rdmto, pro_lineadet.prioridad, pro_tareas.descripcion, pro_lineadet.kghora, pro_lineadet.numop, pro_lineadet.procesado, pro_lineadet.numopideal, pro_lineadet.eficop, pro_lineadet.efictar, pro_lineadet.cantreal, pro_lineadet.durtar, pro_lineadet.factor, pro_lineadet.desvalance, pro_lineadet.intervalo, pro_lineadet.durtarreal
'FROM pro_lineadet LEFT JOIN pro_tareas ON pro_lineadet.idtar = pro_tareas.id
'Where (((pro_lineadet.idlineadet) = 4))
'GROUP BY pro_lineadet.idlineadet, pro_lineadet.orden, pro_lineadet.idrec, pro_lineadet.idtar, pro_lineadet.idunimed, pro_lineadet.rdmto, pro_lineadet.prioridad, pro_tareas.descripcion, pro_lineadet.kghora, pro_lineadet.numop, pro_lineadet.procesado, pro_lineadet.numopideal, pro_lineadet.eficop, pro_lineadet.efictar, pro_lineadet.cantreal, pro_lineadet.durtar, pro_lineadet.factor, pro_lineadet.desvalance, pro_lineadet.intervalo, pro_lineadet.durtarreal
'ORDER BY pro_lineadet.idlineadet, pro_lineadet.orden;

    cSQL = "SELECT pro_lineadet.idlineadet, pro_lineadet.orden, pro_lineadet.idrec, pro_lineadet.idtar, pro_lineadet.idunimed, pro_lineadet.rdmto, pro_lineadet.prioridad, pro_tareas.descripcion, pro_lineadet.kghora, pro_lineadet.numop, pro_lineadet.procesado, pro_lineadet.numopideal, pro_lineadet.eficop, pro_lineadet.efictar, pro_lineadet.cantreal, pro_lineadet.durtar, pro_lineadet.factor, pro_lineadet.desvalance, pro_lineadet.intervalo, pro_lineadet.durtarreal " _
        + vbCr + "FROM pro_lineadet LEFT JOIN pro_tareas ON pro_lineadet.idtar = pro_tareas.id " _
        + vbCr + "Where (((pro_lineadet.idlineadet) = " & idCodigo & ")) " _
        + vbCr + "GROUP BY pro_lineadet.idlineadet, pro_lineadet.orden, pro_lineadet.idrec, pro_lineadet.idtar, pro_lineadet.idunimed, pro_lineadet.rdmto, pro_lineadet.prioridad, pro_tareas.descripcion, pro_lineadet.kghora, pro_lineadet.numop, pro_lineadet.procesado, pro_lineadet.numopideal, pro_lineadet.eficop, pro_lineadet.efictar, pro_lineadet.cantreal, pro_lineadet.durtar, pro_lineadet.factor, pro_lineadet.desvalance, pro_lineadet.intervalo, pro_lineadet.durtarreal " _
        + vbCr + "ORDER BY pro_lineadet.idlineadet, pro_lineadet.orden;"
            
    RST_Busq RstTmp, cSQL, xCon
    
    If RstValores.State = 0 Then DEFINIR_RST_TMP RstValores, RstTmp
    CARGAR_RST_TMP RstValores, RstTmp
    
    Set RstTmp = Nothing
End Sub

Private Sub limpiarRST(Rst As ADODB.Recordset, Optional TODO As Boolean = True)
    With Rst
        If .State <> 0 Then
            If TODO Then .Filter = adFilterNone
            If .RecordCount <> 0 Then
                .MoveFirst
                While Not .EOF
                    .Delete
                    .MoveNext
                Wend
            End If
        End If
    End With
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarDatosValores
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL DETALLE DEL REGISTRO ACTUAL
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarDatosValores()
    Agregando = True
    
    If RstValores.State = 0 Then Exit Sub
    If RstValores.RecordCount = 0 Then Exit Sub
    
    RstValores.MoveFirst
    
    With fg(1)
        .Rows = 1
        Do While Not RstValores.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = NulosC(RstValores("orden"))
            .TextMatrix(.Rows - 1, 2) = NulosC(RstValores("descripcion"))
            .TextMatrix(.Rows - 1, 3) = NulosC(RstValores("rdmto"))
            .TextMatrix(.Rows - 1, 5) = NulosC(RstValores("kghora"))
            .TextMatrix(.Rows - 1, 7) = NulosC(RstValores("prioridad"))
            .TextMatrix(.Rows - 1, 8) = NulosN(RstValores("numop"))
            
            .TextMatrix(.Rows - 1, 4) = NulosN(RstValores("procesado"))
            .TextMatrix(.Rows - 1, 6) = NulosN(RstValores("numopideal"))
            .TextMatrix(.Rows - 1, 9) = NulosN(RstValores("eficop"))
            .TextMatrix(.Rows - 1, 10) = NulosN(RstValores("efictar"))
            .TextMatrix(.Rows - 1, 11) = NulosN(RstValores("cantreal"))
            .TextMatrix(.Rows - 1, 12) = NulosN(RstValores("durtar"))
            .TextMatrix(.Rows - 1, 13) = NulosN(RstValores("factor"))
            .TextMatrix(.Rows - 1, 14) = NulosN(RstValores("desvalance"))
            .TextMatrix(.Rows - 1, 15) = NulosN(RstValores("intervalo"))
            .TextMatrix(.Rows - 1, 16) = NulosN(RstValores("durtarreal"))
            
            .TextMatrix(.Rows - 1, 17) = NulosN(RstValores("idlineadet"))
            .TextMatrix(.Rows - 1, 18) = NulosN(RstValores("idtar"))
            .TextMatrix(.Rows - 1, 19) = NulosN(RstValores("idunimed"))
            
            RstValores.MoveNext
        Loop
    End With
    
    Agregando = False
End Sub

'*****************************************************************************************************
'* Nombre           : pRegistroAdd
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : AÑADE UNA FILA AL CONTROL Fg
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
'Private Sub pRegistroAdd()
'    Dim mCol%
'    Dim fInsertar As Boolean
'    Agregando = True
'    If Fg(0).Row < 1 Then
'        MsgBox "Falta especificar la Tarea", vbInformation, xTitulo
'        Exit Sub
'    End If
'    If NulosN(Fg(0).TextMatrix(Fg(0).Row, 2)) = 0 Then
'        MsgBox "Falta especificar la Tarea", vbInformation, xTitulo
'        Exit Sub
'    End If
'    If Fg(1).Rows > Fg(1).FixedRows Then
'        If NulosC(Fg(1).TextMatrix(Fg(1).Rows - 1, 1)) = "" Then    ' descripcion de unidad
'            MsgBox "Seleccione la Unidad", vbInformation, xTitulo
'        Else
'            fInsertar = True
'        End If
'    Else
'        fInsertar = True
'    End If
'    mCol = 2
'
'    If fInsertar = True Then Fg(1).AddItem ""
'    Fg(1).Row = Fg(1).Rows - 1
'    Fg(1).Col = mCol
'    If fInsertar = True Then Fg_CellButtonClick 1, Fg(1).Rows - 1, 1
'    Fg(1).SetFocus
'    Agregando = False
'End Sub

'*****************************************************************************************************
'* Nombre           : pRegistroDel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UNA FILA DEL CONTROL Fg
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
'Private Sub pRegistroDel()
'    If Fg(1).Row < 1 Then
'        MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Fg(1).SetFocus
'        Exit Sub
'    End If
'
'    If Fg(1).Rows = 1 Then
'        MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Fg(1).SetFocus
'        Exit Sub
'    End If
'    If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
'
'    ' aplicando filtro
'    RstValores.Filter = "idtar = " & Fg(0).TextMatrix(Fg(0).Row, 2)
'    If RstValores.RecordCount <> 0 Then RstValores.MoveFirst
'    RstValores.Find "idunimed  = " & NulosN(Fg(1).TextMatrix(Fg(1).Row, 10))
'    If RstValores.EOF = False And RstValores.BOF = False Then
'        RstValores.Delete
'    End If
'    Fg(1).RemoveItem Fg(1).Row
'
'    If Fg(1).Rows > 1 Then
'        Fg(1).Row = Fg(1).Rows - 1
'        Fg(1).Col = 1
'        Fg(1).SetFocus
'    Else
'        Cmd(0).SetFocus
'    End If
'End Sub

'Private Sub pDatosCostoCargar()
'        Dim Rst As New ADODB.Recordset
'        Dim nSQL As String
'        Dim nSQLTarea As String
'        If NulosN(txt_cb1(0).Text) <> 0 Then
'            nSQLTarea = " and pro_costodet.idtar = " & NulosN(txt_cb1(0).Text)
'        End If
'
'        ' cargando datos de las tareas directas
'        nSQL = "SELECT pro_costodet.idcos, pro_costodet.corr, pro_costo.idref AS idrec, pro_costodet.idtar, pro_costodet.idunimed, alm_inventario.descripcion AS producto, pro_receta.codrec, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.cant, pro_costodet.costo, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio " _
'            + vbCr + " FROM (alm_inventario RIGHT JOIN (pro_receta RIGHT JOIN pro_costo ON pro_receta.id = pro_costo.idref) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (mae_unidades INNER JOIN (pro_tareas INNER JOIN pro_costodet ON pro_tareas.id = pro_costodet.idtar) ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos " _
'            + vbCr + " Where (((pro_costodet.idunimed) <> 7) And ((pro_costo.Tipo) = 1) And ((pro_costo.activo) = -1)) " & nSQLTarea _
'            + vbCr + " ORDER BY alm_inventario.descripcion, pro_receta.codrec, pro_costo.tipo; "
'
'        RST_Busq Rst, nSQL, xCon
'        Agregando = True
'
'        fg1(0).Rows = 1
'        If Rst.RecordCount <> 0 Then
'            Do While Not Rst.EOF
'                fg1(0).Rows = fg1(0).Rows + 1
'                fg1(0).TextMatrix(fg1(0).Rows - 1, 1) = NulosC(Rst("producto"))
'                fg1(0).TextMatrix(fg1(0).Rows - 1, 2) = NulosC(Rst("codrec"))
'                fg1(0).TextMatrix(fg1(0).Rows - 1, 3) = NulosC(Rst("tarea"))
'                fg1(0).TextMatrix(fg1(0).Rows - 1, 4) = NulosC(Rst("abrev"))
'                fg1(0).TextMatrix(fg1(0).Rows - 1, 5) = NulosN(Rst("cant"))
'                fg1(0).TextMatrix(fg1(0).Rows - 1, 6) = NulosN(Rst("costo"))
'                fg1(0).TextMatrix(fg1(0).Rows - 1, 7) = NulosN(Rst("idrec"))
'                fg1(0).TextMatrix(fg1(0).Rows - 1, 8) = NulosN(Rst("idtar"))
'                fg1(0).TextMatrix(fg1(0).Rows - 1, 9) = NulosN(Rst("idunimed"))
'                fg1(0).TextMatrix(fg1(0).Rows - 1, 10) = NulosN(Rst("idcos"))
'                fg1(0).TextMatrix(fg1(0).Rows - 1, 11) = NulosN(Rst("corr"))
'                If NulosN(Rst("costo")) = 0 Then
'                    GRID_COLOR_FONDO fg1(0), fg1(0).Rows - 1, 6, fg1(0).Rows - 1, 6, vbRed
'                End If
'                Rst.MoveNext
'            Loop
'        End If
'        Set Rst = Nothing
'
'        nSQL = "SELECT pro_costodet.idcos,pro_costodet.corr, pro_costodet.idtar, pro_costodet.idunimed, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.cant, pro_costodet.costo, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio " _
'            + vbCr + " FROM pro_tareas INNER JOIN (pro_costo INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costo.idref " _
'            + vbCr + " Where (((pro_costodet.idunimed) <> 7) And ((pro_costo.Tipo) = 2) And ((pro_costo.activo) = -1)) " & nSQLTarea _
'            + vbCr + " ORDER BY pro_tareas.descripcion, pro_costo.tipo; "
'
'        RST_Busq Rst, nSQL, xCon
'
'        fg1(1).Rows = 1
'        If Rst.RecordCount <> 0 Then
'            Do While Not Rst.EOF
'                fg1(1).Rows = fg1(1).Rows + 1
'                fg1(1).TextMatrix(fg1(1).Rows - 1, 1) = NulosC(Rst("tarea"))
'                fg1(1).TextMatrix(fg1(1).Rows - 1, 2) = NulosC(Rst("abrev"))
'                fg1(1).TextMatrix(fg1(1).Rows - 1, 3) = NulosN(Rst("cant"))
'                fg1(1).TextMatrix(fg1(1).Rows - 1, 4) = NulosN(Rst("costo"))
'                fg1(1).TextMatrix(fg1(1).Rows - 1, 5) = NulosN(Rst("idtar"))
'                fg1(1).TextMatrix(fg1(1).Rows - 1, 6) = NulosN(Rst("idunimed"))
'                fg1(1).TextMatrix(fg1(1).Rows - 1, 7) = NulosN(Rst("idcos"))
'                fg1(1).TextMatrix(fg1(1).Rows - 1, 8) = NulosN(Rst("corr"))
'
'                If NulosN(Rst("costo")) = 0 Then
'                    GRID_COLOR_FONDO fg1(1), fg1(1).Rows - 1, 4, fg1(1).Rows - 1, 4, vbRed
'                End If
'                Rst.MoveNext
'            Loop
'        End If
'        Set Rst = Nothing
'        Agregando = False
'End Sub

'Private Sub pDatosCostoGrabar()
'    Dim mRow&
'    If TabOne2.CurrTab = 0 Then
'        If fg1(0).Rows = 1 Then
'            MsgBox "No hay Lista de tareas con Productos", vbExclamation, xTitulo
'            Exit Sub
'        End If
'    Else
'        If fg1(1).Rows = 1 Then
'            MsgBox "No hay Lista de tareas Diversos", vbExclamation, xTitulo
'            Exit Sub
'        End If
'    End If
'
'    If MsgBox("Seguro desea continuar", vbQuestion + vbYesNo) = vbNo Then Exit Sub
'    On Error GoTo error
'    xCon.BeginTrans
'    If TabOne2.CurrTab = 0 Then
'        For mRow = 1 To fg1(0).Rows - 1
'            xCon.Execute "update pro_costodet set cant= " & NulosN(fg1(0).TextMatrix(mRow, 5)) & ", costo = " & NulosN(fg1(0).TextMatrix(mRow, 6)) & " where idcos= " & NulosN(fg1(0).TextMatrix(mRow, 10)) & "  and corr = " & NulosN(fg1(0).TextMatrix(mRow, 11)) & " and idtar = " & NulosN(fg1(0).TextMatrix(mRow, 8)) & " and idunimed = " & NulosN(fg1(0).TextMatrix(mRow, 9))
'        Next mRow
'    Else
'        For mRow = 1 To fg1(1).Rows - 1
'            xCon.Execute "update pro_costodet set cant= " & NulosN(fg1(1).TextMatrix(mRow, 3)) & ", costo = " & NulosN(fg1(1).TextMatrix(mRow, 4)) & " where idcos= " & NulosN(fg1(1).TextMatrix(mRow, 7)) & "  and corr = " & NulosN(fg1(1).TextMatrix(mRow, 8)) & " and idtar = " & NulosN(fg1(1).TextMatrix(mRow, 5)) & " and idunimed = " & NulosN(fg1(1).TextMatrix(mRow, 6))
'        Next mRow
'    End If
'
'    xCon.CommitTrans
'    MsgBox "Se grabaron las tareas " & IIf(TabOne2.CurrTab = 0, " Directas", " Diversas ") & " con éxito", vbInformation, xTitulo
'
'    Exit Sub
'
'error:
'    xCon.RollbackTrans
'    SHOW_ERROR Me.Name, "pDatosCostoGrabar"
'End Sub
'
'Private Sub pDatosCostoCalcular(Tipo As Integer)
'    ' tipo =1 actualizar costo
'    ' tipo =2 equivalencia de costo a hora
'
'    Dim mRow&
'    If QueHace <> 3 Then Exit Sub
'    If Tipo = 1 Then
'        If NulosN(TxtCosto2.Text) = 0 Then
'            MsgBox "Ingrese el costo por Hora", vbExclamation, xTitulo
'            TxtCosto2.SetFocus
'            Exit Sub
'        End If
'    Else
'        If NulosN(TxtCosto3.Text) = 0 Then
'            MsgBox "Ingrese el costo por Hora", vbExclamation, xTitulo
'            TxtCosto3.SetFocus
'            Exit Sub
'        End If
'    End If
'
'    ' actualizar el costo/unidad de acuerdo al pago por hora
'    Agregando = True
'    If Tipo = 1 Then
'        If TabOne2.CurrTab = 0 Then              ' costo directo
'            For mRow = 0 To fg1(0).Rows - 1
'                If NulosN(fg1(0).TextMatrix(mRow, 5)) <> 0 Then
'                    fg1(0).TextMatrix(mRow, 6) = NulosN(TxtCosto2.Text) / NulosN(fg1(0).TextMatrix(mRow, 5))
'                End If
'            Next mRow
'        Else
'            For mRow = 0 To fg1(1).Rows - 1
'                If NulosN(fg1(1).TextMatrix(mRow, 3)) <> 0 Then
'                    fg1(1).TextMatrix(mRow, 4) = NulosN(TxtCosto2.Text) / NulosN(fg1(1).TextMatrix(mRow, 3))
'                End If
'            Next mRow
'        End If
'    Else
'        For mRow = 0 To fg1(2).Rows - 1
'            If NulosN(fg1(2).TextMatrix(mRow, 5)) <> 0 Then
'                fg1(2).TextMatrix(mRow, 6) = NulosN(TxtCosto3.Text) / NulosN(fg1(2).TextMatrix(mRow, 5))
'            End If
'        Next mRow
'    End If
'    Agregando = False
'
'End Sub

Private Sub PbCerrar_Click(Index As Integer)
    If Index = 0 Then
        Frame4.Visible = False
    End If
End Sub

'Private Sub pConvertAHoraCargar()
'    Dim Rst  As New ADODB.Recordset
'    Dim nSQL As String
'
'    If TxtFecha(0).valor = "" Or TxtFecha(1).valor = "" Then
'        MsgBox "Ingrese una fecha", vbExclamation, xTitulo
'        If TxtFecha(0).valor = "" Then TxtFecha(0).SetFocus Else TxtFecha(1).SetFocus
'        Exit Sub
'    End If
'    If CDate(TxtFecha(0).valor) > CDate(TxtFecha(1).valor) Then
'        MsgBox "La fecha inicial es superior al Final", vbExclamation, xTitulo
'        TxtFecha(0).SetFocus
'        Exit Sub
'    End If
'
'    nSQL = "SELECT vwTarea.codigopk, vwTarea.tarea, vwTarea.producto, vwTarea.abrev, vwTarea.idtar, vwTarea.idrec, vwTarea.idunimed, vwcosto.canteo, vwcosto.costo,vwcosto.idcos,vwcosto.corr,vwcosto.paghor " _
'        + vbCr + " FROM ( "
'    nSQL = nSQL _
'        + vbCr + " SELECT  IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk,pro_controltar.idarea, pro_controltardet.idtar, pro_controltardet.idrec, pro_controltardet.idunimed, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, mae_unidades.abrev " _
'        + vbCr + " FROM pro_controltar INNER JOIN (alm_inventario RIGHT JOIN (((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr " _
'        + vbCr + " WHERE (((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltar.tipo)=2) AND ((pro_controltardet.tipo)=1)) "
'    nSQL = nSQL _
'        + vbCr + " UNION "
'    nSQL = nSQL _
'        + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk,pro_controltar.idarea, pro_controltardet.idtar, pro_controltardet.idrec, pro_controltardet.idunimed, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, mae_unidades.abrev " _
'        + vbCr + " FROM pro_controltar INNER JOIN ((alm_inventario RIGHT JOIN (((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN pro_controltardetgr ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) ON pro_controltar.id = pro_controltardet.idctr " _
'        + vbCr + " WHERE (((pro_controltardet.idtar)<>0) AND ((pro_controltar.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "')) AND ((pro_controltar.tipo)=2) AND ((pro_controltardet.tipo)=2) AND ((pro_controltardetgr.activo)=-1)) "
'    nSQL = nSQL _
'        + vbCr + " ) AS vwtarea "
'    nSQL = nSQL _
'        + vbCr + " Left Join " _
'        + vbCr + " (SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.idcos,pro_costodet.corr,pro_costodet.paghor,pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden  " _
'        + vbCr + "  FROM pro_tareas INNER JOIN ((alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
'        + vbCr + " ) AS vwcosto"
'    nSQL = nSQL _
'        + vbCr + " ON vwtarea.codigopk = vwcosto.codigopk "
'
'    nSQL = nSQL _
'     + vbCr + " WHERE (((vwTarea.tarea) Is Not Null) AND ((vwTarea.idunimed)<>7)) " _
'     + vbCr + " ORDER BY vwTarea.producto, vwTarea.tarea; "
'
'    RST_Busq Rst, nSQL, xCon
'
'    Agregando = True
'
'    fg1(2).Rows = 1
'    If Rst.RecordCount <> 0 Then
'        Do While Not Rst.EOF
'            fg1(2).Rows = fg1(2).Rows + 1
'            fg1(2).TextMatrix(fg1(2).Rows - 1, 1) = NulosC(Rst("tarea"))
'            fg1(2).TextMatrix(fg1(2).Rows - 1, 2) = NulosC(Rst("producto"))
''            Fg1(2).TextMatrix(Fg1(2).Rows - 1, 3) = NulosC(rst("codrec"))
'            fg1(2).TextMatrix(fg1(2).Rows - 1, 4) = NulosC(Rst("abrev"))
'            fg1(2).TextMatrix(fg1(2).Rows - 1, 5) = NulosN(Rst("canteo"))
'            fg1(2).TextMatrix(fg1(2).Rows - 1, 6) = NulosN(Rst("costo"))
'            fg1(2).TextMatrix(fg1(2).Rows - 1, 7) = NulosN(Rst("idrec"))
'            fg1(2).TextMatrix(fg1(2).Rows - 1, 8) = NulosN(Rst("idtar"))
'            fg1(2).TextMatrix(fg1(2).Rows - 1, 9) = NulosN(Rst("idunimed"))
'            fg1(2).TextMatrix(fg1(2).Rows - 1, 10) = NulosN(Rst("idcos"))
'            fg1(2).TextMatrix(fg1(2).Rows - 1, 11) = NulosN(Rst("corr"))
'            fg1(2).TextMatrix(fg1(2).Rows - 1, 12) = NulosN(Rst("paghor"))
'
'            If NulosN(Rst("costo")) = 0 Then
'                GRID_COLOR_FONDO fg1(2), fg1(2).Rows - 1, 6, fg1(2).Rows - 1, 6, vbRed
'            End If
'
'            Rst.MoveNext
'        Loop
'    End If
'    Set Rst = Nothing
'
'    Agregando = False
'End Sub
'
'Private Sub pConvertAHoraGrabar()
'    Dim mRow&
'
'    If fg1(2).Rows = 1 Then
'        MsgBox "No hay Lista de tareas con Productos", vbExclamation, xTitulo
'        TxtFecha(0).SetFocus
'        Exit Sub
'    End If
'
'    If MsgBox("Seguro desea continuar", vbQuestion + vbYesNo) = vbNo Then Exit Sub
'    On Error GoTo error
'    xCon.BeginTrans
'
'    For mRow = 1 To fg1(2).Rows - 1
'        xCon.Execute "update pro_costodet set cant= " & NulosN(fg1(2).TextMatrix(mRow, 5)) & ", costo = " & NulosN(fg1(2).TextMatrix(mRow, 6)) & ", paghor= " & NulosN(fg1(2).TextMatrix(mRow, 12)) & " where idcos= " & NulosN(fg1(2).TextMatrix(mRow, 10)) & "  and corr = " & NulosN(fg1(2).TextMatrix(mRow, 11)) & " and idtar = " & NulosN(fg1(2).TextMatrix(mRow, 8)) & " and idunimed = " & NulosN(fg1(2).TextMatrix(mRow, 9))
'    Next mRow
'
'    xCon.CommitTrans
'    MsgBox "Se grabaron las tareas con éxito", vbInformation, xTitulo
'    Exit Sub
'
'error:
'    xCon.RollbackTrans
'    SHOW_ERROR Me.Name, "pDatosCostoGrabar"
'End Sub

'Metodos para arrastrar el Frame
''''''''''''''''''''''''''''''''
Private Sub Frame4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    Frame4.ZOrder 0
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With Frame4
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub
