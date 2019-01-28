VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#12.0#0"; "CODEJO~1.OCX"
Begin VB.Form FrmCronoProduccion2_1 
   Caption         =   "Produccion - Cronograma de Produccion"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   13200
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   480
      TabIndex        =   147
      Top             =   7620
      Visible         =   0   'False
      Width           =   4740
      Begin VB.Shape Shape1 
         Height          =   765
         Index           =   3
         Left            =   60
         Top             =   90
         Width           =   4605
      End
      Begin VB.Label LblProg 
         AutoSize        =   -1  'True
         Caption         =   "CONTROL DE REGISTROS"
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
         Left            =   1920
         TabIndex        =   150
         Top             =   180
         Width           =   2025
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
         Left            =   435
         TabIndex        =   149
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ESPERE POR FAVOR ..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   1470
         TabIndex        =   148
         Top             =   480
         Width           =   1770
      End
   End
   Begin VB.Frame FrmAdd 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5760
      Left            =   2190
      TabIndex        =   65
      Top             =   1350
      Visible         =   0   'False
      Width           =   8880
      Begin VB.CommandButton Cmd 
         Enabled         =   0   'False
         Height          =   240
         Index           =   16
         Left            =   2130
         Picture         =   "FrmCronoProduccion2.1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Agregar Producto"
         Top             =   1110
         Width           =   225
      End
      Begin VB.TextBox TxtNumProd 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   300
         Left            =   1980
         TabIndex        =   1
         Text            =   "TxtNumProd"
         Top             =   360
         Width           =   6795
      End
      Begin SizerOneLibCtl.TabOne TabOne2 
         Height          =   3795
         Left            =   90
         TabIndex        =   78
         Top             =   1425
         Width           =   8670
         _cx             =   15293
         _cy             =   6685
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
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
         ForeColor       =   0
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483637
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   8388608
         Caption         =   " &Tareas | &Personal"
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
         TabCaptionPos   =   3
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         Begin VB.Frame Frame8 
            BorderStyle     =   0  'None
            Height          =   3705
            Left            =   9615
            TabIndex        =   80
            Top             =   45
            Width           =   8280
            Begin VB.Frame Frame14 
               Height          =   555
               Left            =   50
               TabIndex        =   90
               Top             =   -80
               Width           =   8190
               Begin VB.OptionButton OptPers 
                  Caption         =   "Tarea"
                  Height          =   285
                  Index           =   1
                  Left            =   1020
                  TabIndex        =   112
                  Top             =   160
                  Width           =   765
               End
               Begin VB.OptionButton OptPers 
                  Caption         =   "Todos"
                  Height          =   285
                  Index           =   0
                  Left            =   90
                  TabIndex        =   111
                  Top             =   160
                  Width           =   825
               End
               Begin VB.CommandButton Cmd 
                  Height          =   240
                  Index           =   3
                  Left            =   2370
                  Picture         =   "FrmCronoProduccion2.1.frx":0132
                  Style           =   1  'Graphical
                  TabIndex        =   91
                  ToolTipText     =   "Seleccionar Tarea"
                  Top             =   180
                  Width           =   240
               End
               Begin VB.TextBox TxtOrden 
                  Height          =   300
                  Left            =   1830
                  MaxLength       =   12
                  TabIndex        =   92
                  Text            =   "TxtOrden"
                  Top             =   150
                  Width           =   810
               End
               Begin VB.Label LblIdTarea 
                  AutoSize        =   -1  'True
                  Caption         =   "LblIdTarea"
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Left            =   7230
                  TabIndex        =   93
                  Top             =   210
                  Visible         =   0   'False
                  Width           =   765
               End
               Begin VB.Label LblTarea 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblTarea"
                  ForeColor       =   &H00800000&
                  Height          =   300
                  Left            =   2640
                  TabIndex        =   94
                  Top             =   150
                  Width           =   5505
               End
            End
            Begin VSFlex7Ctl.VSFlexGrid fg 
               Height          =   2400
               Index           =   1
               Left            =   45
               TabIndex        =   83
               Top             =   510
               Width           =   8160
               _cx             =   14393
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
               Rows            =   1
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmCronoProduccion2.1.frx":0264
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
            Begin VB.Frame Frame13 
               Height          =   555
               Left            =   50
               TabIndex        =   84
               Top             =   3150
               Width           =   8190
               Begin VB.CommandButton Cmd 
                  Caption         =   "Ver Ranking"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   8
                  Left            =   6840
                  TabIndex        =   89
                  ToolTipText     =   "Mostrar Ranking de Personal para la Tarea Seleccionada"
                  Top             =   150
                  Width           =   1200
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "Agregar"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   4
                  Left            =   60
                  TabIndex        =   88
                  ToolTipText     =   "Agregar Personal"
                  Top             =   150
                  Width           =   1155
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "&Eliminar"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   6
                  Left            =   2760
                  TabIndex        =   87
                  TabStop         =   0   'False
                  ToolTipText     =   "Eliminar Personal"
                  Top             =   150
                  Width           =   1155
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "&Seleccionar"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   5
                  Left            =   1260
                  TabIndex        =   86
                  TabStop         =   0   'False
                  ToolTipText     =   "Agregar Personal de una Lista"
                  Top             =   150
                  Width           =   1155
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "Eliminar Todos"
                  Enabled         =   0   'False
                  Height          =   330
                  Index           =   7
                  Left            =   3960
                  TabIndex        =   85
                  ToolTipText     =   "Eliminar Todos"
                  Top             =   150
                  Width           =   1200
               End
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Nº Trabajadores"
               Height          =   195
               Left            =   5910
               TabIndex        =   96
               Top             =   2960
               Width           =   1155
            End
            Begin VB.Label LblDetTrab 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "LbDetTrab"
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
               Left            =   7230
               TabIndex        =   95
               Top             =   2960
               Width           =   915
            End
         End
         Begin VB.Frame Frame6 
            BorderStyle     =   0  'None
            Height          =   3705
            Left            =   345
            TabIndex        =   79
            Top             =   45
            Width           =   8280
            Begin VB.TextBox TxtCantMP 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   3082
                  SubFormatType   =   1
               EndProperty
               Height          =   315
               Left            =   1890
               Locked          =   -1  'True
               TabIndex        =   136
               Text            =   "TxtCantMP"
               Top             =   2880
               Width           =   1005
            End
            Begin VSFlex7Ctl.VSFlexGrid fg 
               Height          =   2310
               Index           =   0
               Left            =   0
               TabIndex        =   82
               Top             =   540
               Width           =   8235
               _cx             =   14526
               _cy             =   4075
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
               Cols            =   13
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmCronoProduccion2.1.frx":0354
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
            Begin VB.Frame Frame19 
               Height          =   555
               Left            =   0
               TabIndex        =   131
               Top             =   3150
               Width           =   8220
               Begin VB.CommandButton Cmd 
                  Enabled         =   0   'False
                  Height          =   240
                  Index           =   20
                  Left            =   5580
                  Picture         =   "FrmCronoProduccion2.1.frx":04C9
                  Style           =   1  'Graphical
                  TabIndex        =   9
                  ToolTipText     =   "Seleccionar Tarea"
                  Top             =   210
                  Width           =   240
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "&Propiedades"
                  Enabled         =   0   'False
                  Height          =   350
                  Index           =   1
                  Left            =   1740
                  TabIndex        =   135
                  ToolTipText     =   "Mostrar Propiedades de Procesado de Tareas"
                  Top             =   150
                  Width           =   1155
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "&Procesar"
                  Enabled         =   0   'False
                  Height          =   350
                  Index           =   2
                  Left            =   2970
                  TabIndex        =   11
                  ToolTipText     =   "Procesar las Tareas del Producto Seleccionado"
                  Top             =   150
                  Width           =   1155
               End
               Begin VB.CommandButton Cmd 
                  Caption         =   "&Linea de Produccion"
                  Height          =   350
                  Index           =   9
                  Left            =   60
                  TabIndex        =   132
                  ToolTipText     =   "Modificar - Editar la Linea de Produccion"
                  Top             =   150
                  Width           =   1635
               End
               Begin VB.TextBox TxtIdLineaDet 
                  Height          =   300
                  Left            =   5070
                  MaxLength       =   12
                  TabIndex        =   10
                  Text            =   "TxtIdLinea"
                  Top             =   180
                  Width           =   780
               End
               Begin VB.Label LbldLineaDet 
                  AutoSize        =   -1  'True
                  Caption         =   "LbldLineaDet"
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Left            =   7020
                  TabIndex        =   144
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   1005
               End
               Begin VB.Label Label23 
                  AutoSize        =   -1  'True
                  Caption         =   "Linea"
                  Height          =   195
                  Left            =   4600
                  TabIndex        =   143
                  Top             =   210
                  Width           =   390
               End
               Begin VB.Label LblLinea 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblLinea"
                  ForeColor       =   &H00800000&
                  Height          =   300
                  Left            =   5880
                  TabIndex        =   142
                  Top             =   180
                  Width           =   2220
               End
            End
            Begin VB.Frame Frame5 
               Height          =   555
               Left            =   0
               TabIndex        =   81
               Top             =   -80
               Width           =   8235
               Begin VB.CommandButton Cmd 
                  Enabled         =   0   'False
                  Height          =   240
                  Index           =   18
                  Left            =   1470
                  Picture         =   "FrmCronoProduccion2.1.frx":05FB
                  Style           =   1  'Graphical
                  TabIndex        =   7
                  ToolTipText     =   "Seleccionar Tarea"
                  Top             =   180
                  Width           =   240
               End
               Begin VB.TextBox TxtIdEncarg 
                  Height          =   300
                  Left            =   960
                  MaxLength       =   12
                  TabIndex        =   8
                  Text            =   "TxtIdEncarg"
                  Top             =   150
                  Width           =   780
               End
               Begin VB.Label LblEncargado 
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "LblEncargado"
                  ForeColor       =   &H00800000&
                  Height          =   300
                  Left            =   1770
                  TabIndex        =   134
                  Top             =   150
                  Width           =   6345
               End
               Begin VB.Label Label26 
                  AutoSize        =   -1  'True
                  Caption         =   "Encargado"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   133
                  Top             =   180
                  Width           =   780
               End
            End
            Begin MSComCtl2.DTPicker DTPHoraFin 
               Height          =   345
               Left            =   7320
               TabIndex        =   137
               Top             =   2880
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   609
               _Version        =   393216
               CustomFormat    =   "HH:mm"
               Format          =   57933827
               UpDown          =   -1  'True
               CurrentDate     =   40606
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtBFchFin 
               Height          =   300
               Left            =   5460
               TabIndex        =   145
               Top             =   2850
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
               Locked          =   -1  'True
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Fch. Fin"
               Height          =   195
               Left            =   4830
               TabIndex        =   146
               Top             =   2880
               Width           =   570
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Insumo Base Requerido"
               Height          =   195
               Left            =   60
               TabIndex        =   139
               Top             =   2910
               Width           =   1695
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Hor. Fin"
               Height          =   195
               Left            =   6720
               TabIndex        =   138
               Top             =   2910
               Width           =   555
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   555
         Left            =   60
         TabIndex        =   67
         Top             =   5145
         Width           =   8715
         Begin VB.CommandButton Cmd 
            Caption         =   "&Imprimir Linea de Produccion"
            Height          =   350
            Index           =   19
            Left            =   60
            TabIndex        =   140
            ToolTipText     =   "Aceptar Edicion del Producto"
            Top             =   130
            Width           =   2445
         End
         Begin VB.CommandButton Cmd 
            Caption         =   "&Cancelar"
            Height          =   350
            Index           =   11
            Left            =   7470
            TabIndex        =   69
            ToolTipText     =   "Cancelar Edicion del Producto"
            Top             =   130
            Width           =   1155
         End
         Begin VB.CommandButton Cmd 
            Caption         =   "&Aceptar"
            Enabled         =   0   'False
            Height          =   350
            Index           =   10
            Left            =   6240
            TabIndex        =   68
            ToolTipText     =   "Aceptar Edicion del Producto"
            Top             =   130
            Width           =   1155
         End
         Begin VB.Label LblNTrab 
            AutoSize        =   -1  'True
            Caption         =   "LblNTrab"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   5340
            TabIndex        =   153
            Top             =   180
            Visible         =   0   'False
            Width           =   660
         End
      End
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   8580
         Picture         =   "FrmCronoProduccion2.1.frx":072D
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   66
         ToolTipText     =   "Cerrar"
         Top             =   60
         Width           =   195
      End
      Begin VB.TextBox TxtCant 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   300
         Left            =   3975
         TabIndex        =   5
         Text            =   "TxtCant"
         Top             =   1080
         Width           =   885
      End
      Begin VB.CommandButton Cmd 
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   1680
         Picture         =   "FrmCronoProduccion2.1.frx":0A19
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Agregar Producto"
         Top             =   750
         Width           =   225
      End
      Begin MSComCtl2.DTPicker DTPHoras 
         Height          =   345
         Left            =   7260
         TabIndex        =   6
         Top             =   1050
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   57933827
         UpDown          =   -1  'True
         CurrentDate     =   40606
      End
      Begin VB.TextBox TxtMatProd 
         Height          =   300
         Left            =   1005
         TabIndex        =   3
         Text            =   "TxtMatProd"
         Top             =   720
         Width           =   915
      End
      Begin VB.TextBox TxtCodRec 
         Height          =   300
         Left            =   1005
         TabIndex        =   4
         Text            =   "TxtCodRec"
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Label lblIdRec 
         AutoSize        =   -1  'True
         Caption         =   "lblIdRec"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   8010
         TabIndex        =   155
         Top             =   780
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Receta"
         Height          =   195
         Left            =   180
         TabIndex        =   154
         Top             =   1140
         Width           =   525
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Numero de Produccion"
         Height          =   195
         Left            =   180
         TabIndex        =   152
         Top             =   400
         Width           =   1635
      End
      Begin VB.Label LblIdCrDet 
         AutoSize        =   -1  'True
         Caption         =   "LblIdCrDet"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   6360
         TabIndex        =   98
         Top             =   60
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label LblNomOperacion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agregando Cronograma"
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
         Left            =   120
         TabIndex        =   76
         Top             =   60
         Width           =   1995
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   15
         X2              =   6045
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   15
         X2              =   0
         Y1              =   0
         Y2              =   5730
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         X1              =   8850
         X2              =   8850
         Y1              =   0
         Y2              =   5730
      End
      Begin VB.Label LblDia 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LblDia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   7920
         TabIndex        =   75
         Top             =   60
         Width           =   555
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "HH:mm"
         Height          =   195
         Left            =   8220
         TabIndex        =   74
         Top             =   1140
         Width           =   525
      End
      Begin VB.Label LblUnidad 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblUnidad"
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
         TabIndex        =   73
         Top             =   1080
         Width           =   1590
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         Height          =   195
         Left            =   180
         TabIndex        =   72
         Top             =   780
         Width           =   645
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   3150
         TabIndex        =   71
         Top             =   1140
         Width           =   630
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Hor. Ini."
         Height          =   195
         Left            =   6630
         TabIndex        =   70
         Top             =   1140
         Width           =   555
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         X1              =   0
         X2              =   8850
         Y1              =   5730
         Y2              =   5730
      End
      Begin VB.Label LblMatProd 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblMatProd"
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
         Left            =   1980
         TabIndex        =   77
         Top             =   720
         Width           =   6780
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Left            =   30
         Top             =   45
         Width           =   8785
      End
   End
   Begin VB.Frame Frame10 
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
      Left            =   6690
      TabIndex        =   113
      Top             =   3570
      Visible         =   0   'False
      Width           =   6330
      Begin VB.CommandButton Cmd 
         Caption         =   "&Cancelar"
         Height          =   350
         Index           =   15
         Left            =   1350
         TabIndex        =   130
         ToolTipText     =   "Cancelar Edicion del Producto"
         Top             =   3900
         Width           =   1155
      End
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   3
         Left            =   6060
         Picture         =   "FrmCronoProduccion2.1.frx":0B4B
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   125
         ToolTipText     =   "Cerrar"
         Top             =   30
         Width           =   195
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "Adicionar"
         Height          =   330
         Index           =   14
         Left            =   100
         TabIndex        =   124
         ToolTipText     =   "Eliminar Todos"
         Top             =   3900
         Width           =   1155
      End
      Begin VB.Frame Frame16 
         Height          =   435
         Left            =   90
         TabIndex        =   121
         Top             =   1110
         Width           =   6120
         Begin VB.OptionButton OptSel 
            Caption         =   "Deselec. Todos"
            Enabled         =   0   'False
            Height          =   225
            Index           =   1
            Left            =   1500
            TabIndex        =   123
            Top             =   150
            Width           =   1485
         End
         Begin VB.OptionButton OptSel 
            Caption         =   "Selec. Todos"
            Enabled         =   0   'False
            Height          =   225
            Index           =   0
            Left            =   60
            TabIndex        =   122
            Top             =   150
            Width           =   1305
         End
      End
      Begin VB.Frame Frame15 
         Height          =   855
         Left            =   90
         TabIndex        =   114
         Top             =   300
         Width           =   6135
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Producto"
            Height          =   195
            Left            =   100
            TabIndex        =   119
            Top             =   135
            Width           =   645
         End
         Begin VB.Label LblProd2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblProd2"
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   840
            TabIndex        =   118
            Top             =   120
            Width           =   5220
         End
         Begin VB.Label LblIdTarea2 
            AutoSize        =   -1  'True
            Caption         =   "LblIdTarea2"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   4260
            TabIndex        =   117
            Top             =   540
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "Tarea"
            Height          =   195
            Index           =   9
            Left            =   100
            TabIndex        =   116
            Top             =   540
            Width           =   420
         End
         Begin VB.Label LblIdprod 
            AutoSize        =   -1  'True
            Caption         =   "LblIdprod"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   4290
            TabIndex        =   115
            Top             =   180
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.Label LblTarea2 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTarea2"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   840
            TabIndex        =   120
            Top             =   480
            Width           =   5205
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   2220
         Index           =   2
         Left            =   90
         TabIndex        =   126
         Top             =   1560
         Width           =   6120
         _cx             =   10795
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCronoProduccion2.1.frx":0E37
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
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   5
         X1              =   6300
         X2              =   6300
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
         X2              =   6300
         Y1              =   4290
         Y2              =   4290
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ranking de Personal"
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
         Left            =   105
         TabIndex        =   129
         Top             =   60
         Width           =   1785
      End
      Begin VB.Label LbNumSel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "LbNumSel"
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
         Left            =   5340
         TabIndex        =   128
         Top             =   3930
         Width           =   870
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Trab. Selec."
         Height          =   195
         Left            =   4260
         TabIndex        =   127
         Top             =   3930
         Width           =   870
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   30
         Top             =   30
         Width           =   6240
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3360
      Left            =   6330
      TabIndex        =   99
      Top             =   3240
      Visible         =   0   'False
      Width           =   6110
      Begin VB.TextBox TxtTotal 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4020
         Locked          =   -1  'True
         TabIndex        =   105
         Text            =   "TxtTotal"
         Top             =   2460
         Width           =   945
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   350
         Index           =   0
         Left            =   3100
         TabIndex        =   104
         Top             =   2865
         Width           =   1155
      End
      Begin VB.CommandButton CmdAcepta 
         Caption         =   "&Aceptar"
         Height          =   350
         Index           =   0
         Left            =   1860
         TabIndex        =   103
         Top             =   2865
         Width           =   1155
      End
      Begin VB.TextBox TxtCan 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1170
         TabIndex        =   102
         Text            =   "TxtCan"
         Top             =   660
         Width           =   1185
      End
      Begin VB.TextBox TxtMP 
         Height          =   300
         Left            =   1170
         TabIndex        =   101
         Text            =   "TxtMP"
         Top             =   360
         Width           =   4845
      End
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   5800
         Picture         =   "FrmCronoProduccion2.1.frx":0F77
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   100
         ToolTipText     =   "Cerrar"
         Top             =   65
         Width           =   195
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg2 
         Height          =   1470
         Left            =   75
         TabIndex        =   106
         Top             =   990
         Width           =   5880
         _cx             =   10372
         _cy             =   2593
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
         BackColorSel    =   -2147483645
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCronoProduccion2.1.frx":1263
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
      Begin VB.Label Label11 
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
         Height          =   165
         Left            =   3090
         TabIndex        =   110
         Top             =   2505
         Width           =   825
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         X1              =   15
         X2              =   6045
         Y1              =   3315
         Y2              =   3315
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   195
         Left            =   75
         TabIndex        =   109
         Top             =   690
         Width           =   630
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   0
         X1              =   6060
         X2              =   6060
         Y1              =   15
         Y2              =   3330
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   15
         X2              =   0
         Y1              =   0
         Y2              =   3315
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   6045
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccion de Productos"
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
         Left            =   120
         TabIndex        =   108
         Top             =   60
         Width           =   2040
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Materia Prima"
         Height          =   195
         Left            =   75
         TabIndex        =   107
         Top             =   390
         Width           =   960
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   30
         Top             =   45
         Width           =   6000
      End
   End
   Begin VB.Frame Frame9 
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
      Height          =   4340
      Left            =   270
      TabIndex        =   40
      Top             =   2430
      Visible         =   0   'False
      Width           =   4980
      Begin VB.CommandButton Cmd 
         Caption         =   "Aceptar"
         Height          =   345
         Index           =   12
         Left            =   2430
         TabIndex        =   62
         Top             =   3870
         Width           =   1155
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "&Cancelar"
         Height          =   345
         Index           =   13
         Left            =   3645
         TabIndex        =   61
         Top             =   3870
         Width           =   1155
      End
      Begin VB.Frame Frame11 
         Caption         =   "Incluir Horas de Refrigerio?"
         Height          =   945
         Left            =   150
         TabIndex        =   52
         Top             =   2800
         Width           =   4660
         Begin VB.OptionButton OptHoras 
            Caption         =   "No"
            Height          =   225
            Index           =   1
            Left            =   1000
            TabIndex        =   54
            Top             =   450
            Width           =   615
         End
         Begin VB.OptionButton OptHoras 
            Caption         =   "Si"
            Height          =   225
            Index           =   0
            Left            =   300
            TabIndex        =   53
            Top             =   450
            Width           =   555
         End
         Begin MSComCtl2.DTPicker DTPHorIni 
            Height          =   345
            Left            =   2700
            TabIndex        =   55
            Top             =   130
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "HH:mm"
            Format          =   57933827
            UpDown          =   -1  'True
            CurrentDate     =   40606
         End
         Begin MSComCtl2.DTPicker DTPHorFin 
            Height          =   345
            Left            =   2700
            TabIndex        =   56
            Top             =   500
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "HH:mm"
            Format          =   57933827
            UpDown          =   -1  'True
            CurrentDate     =   40606
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "( Fin"
            Height          =   195
            Index           =   8
            Left            =   2100
            TabIndex        =   60
            Top             =   585
            Width           =   300
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "HH:mm )"
            Height          =   195
            Index           =   7
            Left            =   3705
            TabIndex        =   59
            Top             =   585
            Width           =   615
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "( Inicio"
            Height          =   195
            Index           =   6
            Left            =   2100
            TabIndex        =   58
            Top             =   225
            Width           =   465
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "HH:mm )"
            Height          =   195
            Index           =   5
            Left            =   3700
            TabIndex        =   57
            Top             =   230
            Width           =   615
         End
      End
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   2
         Left            =   4680
         Picture         =   "FrmCronoProduccion2.1.frx":12FE
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   51
         ToolTipText     =   "Cerrar"
         Top             =   60
         Width           =   195
      End
      Begin VB.Frame Frame12 
         Caption         =   "La tarea Empieza al : "
         Height          =   2385
         Left            =   150
         TabIndex        =   41
         Top             =   300
         Width           =   4660
         Begin VB.OptionButton OptTarea 
            Caption         =   "Finalizar la tarea anterior"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   46
            Top             =   270
            Width           =   2775
         End
         Begin VB.OptionButton OptTarea 
            Caption         =   "Transcurrir un porcentaje de la tarea anterior"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   45
            Top             =   600
            Width           =   4425
         End
         Begin VB.TextBox TxtPctje 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            Height          =   300
            Left            =   2145
            MaxLength       =   12
            TabIndex        =   44
            Text            =   "TxtPctje"
            Top             =   885
            Width           =   840
         End
         Begin VB.OptionButton OptTarea 
            Caption         =   "Transcurrido los minutos de la tarea anterior"
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   43
            Top             =   1320
            Width           =   3855
         End
         Begin VB.OptionButton OptTarea 
            Caption         =   "Segun Linea"
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   42
            Top             =   2040
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker DTPMinutos 
            Height          =   345
            Left            =   2160
            TabIndex        =   64
            Top             =   1590
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "HH:mm"
            Format          =   57933827
            UpDown          =   -1  'True
            CurrentDate     =   40606
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje"
            Height          =   195
            Index           =   1
            Left            =   1245
            TabIndex        =   50
            Top             =   930
            Width           =   765
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   2
            Left            =   3075
            TabIndex        =   49
            Top             =   930
            Width           =   120
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "Minutos"
            Height          =   195
            Index           =   3
            Left            =   1245
            TabIndex        =   48
            Top             =   1620
            Width           =   555
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "HH:mm"
            Height          =   195
            Index           =   4
            Left            =   3075
            TabIndex        =   47
            Top             =   1620
            Width           =   525
         End
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opciones de Procesado de Tareas"
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
         Left            =   105
         TabIndex        =   63
         Top             =   50
         Width           =   2955
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   2
         X1              =   0
         X2              =   4950
         Y1              =   4300
         Y2              =   4300
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   3
         X1              =   30
         X2              =   30
         Y1              =   210
         Y2              =   3315
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   2
         X1              =   4950
         X2              =   4950
         Y1              =   0
         Y2              =   4300
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   3
         X1              =   0
         X2              =   8295
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   40
         Top             =   30
         Width           =   4860
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7125
      Left            =   15
      TabIndex        =   13
      Top             =   360
      Width           =   11850
      _cx             =   20902
      _cy             =   12568
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
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
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "  &Consulta  |   &Detalle   "
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
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   6690
         Left            =   12495
         TabIndex        =   15
         Top             =   390
         Width           =   11760
         Begin VB.CommandButton CmdOpciones 
            Caption         =   "&Cambiar Vista"
            Height          =   400
            Index           =   5
            Left            =   9090
            TabIndex        =   151
            Top             =   330
            Width           =   2200
         End
         Begin VB.Frame Frame2 
            Height          =   1245
            Left            =   0
            TabIndex        =   16
            Top             =   245
            Width           =   9060
            Begin VB.ComboBox ComboSemanas 
               Height          =   315
               ItemData        =   "FrmCronoProduccion2.1.frx":15EA
               Left            =   1020
               List            =   "FrmCronoProduccion2.1.frx":15EC
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   450
               Width           =   1000
            End
            Begin VB.CommandButton CmdBusTip 
               Enabled         =   0   'False
               Height          =   240
               Left            =   1740
               Picture         =   "FrmCronoProduccion2.1.frx":15EE
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   810
               Width           =   255
            End
            Begin VB.CommandButton CmdBusSup 
               Enabled         =   0   'False
               Height          =   240
               Left            =   1740
               Picture         =   "FrmCronoProduccion2.1.frx":1720
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   180
               Width           =   255
            End
            Begin VB.TextBox TxtIdSup 
               Height          =   300
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   18
               Text            =   "TxtIdSup"
               Top             =   150
               Width           =   1000
            End
            Begin VB.TextBox TxtTipPro 
               Height          =   300
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   17
               Text            =   "TxtTipPro"
               Top             =   780
               Width           =   1000
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
               Height          =   300
               Left            =   5415
               TabIndex        =   22
               Top             =   450
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
               Locked          =   -1  'True
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
               Height          =   300
               Left            =   3000
               TabIndex        =   23
               Top             =   450
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
               Locked          =   -1  'True
            End
            Begin VB.Label LabelSemana 
               AutoSize        =   -1  'True
               Caption         =   "Semana"
               Height          =   195
               Left            =   60
               TabIndex        =   30
               Top             =   510
               Width           =   585
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Prod."
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   29
               Top             =   840
               Width           =   735
            End
            Begin VB.Label LblTipoProd 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblTipoProd"
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
               Left            =   2055
               TabIndex        =   28
               Top             =   780
               Width           =   6795
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Supervisor"
               Height          =   195
               Left            =   60
               TabIndex        =   27
               Top             =   195
               Width           =   750
            End
            Begin VB.Label LblSupervisor 
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblSupervisor"
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
               Left            =   2055
               TabIndex        =   26
               Top             =   150
               Width           =   6795
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Fch. Final"
               Height          =   195
               Left            =   4530
               TabIndex        =   25
               Top             =   510
               Width           =   690
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Fch. Inicio"
               Height          =   195
               Left            =   2055
               TabIndex        =   24
               Top             =   510
               Width           =   735
            End
         End
         Begin XtremeCalendarControl.CalendarControl CalendarControl1 
            Height          =   4215
            Left            =   0
            TabIndex        =   31
            Top             =   1860
            Width           =   11715
            _Version        =   786432
            _ExtentX        =   20664
            _ExtentY        =   7435
            _StockProps     =   64
            ViewType        =   3
         End
         Begin VB.CommandButton CmdOpciones 
            Caption         =   "&Procesar"
            Enabled         =   0   'False
            Height          =   400
            Index           =   0
            Left            =   9090
            TabIndex        =   35
            Top             =   1070
            Width           =   2200
         End
         Begin VB.Frame FrmBotones 
            Height          =   585
            Left            =   0
            TabIndex        =   36
            Top             =   6100
            Width           =   11715
            Begin VB.CommandButton CmdOpciones 
               Caption         =   "Reporte de Planeacion"
               Enabled         =   0   'False
               Height          =   330
               Index           =   4
               Left            =   4290
               TabIndex        =   141
               Top             =   150
               Width           =   2085
            End
            Begin VB.CommandButton CmdOpciones 
               Caption         =   "Eliminar"
               Enabled         =   0   'False
               Height          =   330
               Index           =   3
               Left            =   2880
               TabIndex        =   39
               Top             =   150
               Width           =   1305
            End
            Begin VB.CommandButton CmdOpciones 
               Caption         =   "&Modificar"
               Enabled         =   0   'False
               Height          =   330
               Index           =   2
               Left            =   1470
               TabIndex        =   38
               Top             =   150
               Width           =   1305
            End
            Begin VB.CommandButton CmdOpciones 
               Caption         =   "&Agregar"
               Enabled         =   0   'False
               Height          =   330
               Index           =   1
               Left            =   60
               TabIndex        =   37
               Top             =   150
               Width           =   1305
            End
            Begin VB.Label LblIdCr 
               AutoSize        =   -1  'True
               Caption         =   "LblIdCr"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   8760
               TabIndex        =   97
               Top             =   240
               Visible         =   0   'False
               Width           =   495
            End
         End
         Begin VB.Shape ShapeFondo 
            BackColor       =   &H80000000&
            BackStyle       =   1  'Opaque
            Height          =   795
            Left            =   0
            Top             =   1500
            Width           =   11715
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Cronograma"
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
            Height          =   315
            Left            =   0
            TabIndex        =   32
            Top             =   -10
            Width           =   11655
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   6690
         Left            =   45
         TabIndex        =   14
         Top             =   390
         Width           =   11760
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6285
            Left            =   30
            TabIndex        =   33
            Top             =   360
            Width           =   11700
            _ExtentX        =   20638
            _ExtentY        =   11086
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Semana"
            Columns(0).DataField=   "semana"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Fch. Ini."
            Columns(1).DataField=   "fchini"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Fin."
            Columns(2).DataField=   "fchfin"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Tipo Produccion"
            Columns(3).DataField=   "destippro"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Programador"
            Columns(4).DataField=   "apenom"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1535"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1455"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2223"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2143"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=2249"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2170"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=3757"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3678"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=9102"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=9022"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Named:id=33:Normal"
            _StyleDefs(57)  =   ":id=33,.parent=0"
            _StyleDefs(58)  =   "Named:id=34:Heading"
            _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(60)  =   ":id=34,.wraptext=-1"
            _StyleDefs(61)  =   "Named:id=35:Footing"
            _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(63)  =   "Named:id=36:Selected"
            _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(65)  =   "Named:id=37:Caption"
            _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(67)  =   "Named:id=38:HighlightRow"
            _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=39:EvenRow"
            _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(71)  =   "Named:id=40:OddRow"
            _StyleDefs(72)  =   ":id=40,.parent=33"
            _StyleDefs(73)  =   "Named:id=41:RecordSelector"
            _StyleDefs(74)  =   ":id=41,.parent=34"
            _StyleDefs(75)  =   "Named:id=42:FilterBar"
            _StyleDefs(76)  =   ":id=42,.parent=33"
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Cronogramas"
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
            Height          =   315
            Left            =   0
            TabIndex        =   34
            Top             =   -10
            Width           =   11700
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.1.frx":1852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.1.frx":1D96
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.1.frx":1F1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.1.frx":236E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.1.frx":2486
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.1.frx":29CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.1.frx":2F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.1.frx":3022
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.1.frx":3136
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.1.frx":358A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion2.1.frx":36F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13200
      _ExtentX        =   23283
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   11
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   11
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Linea de Produccion"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Text            =   "Linea de Acabado"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Text            =   "Reporte de Cronograma"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menu_01 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu_01_02 
         Caption         =   "Agregar Producto"
      End
      Begin VB.Menu menu_01_04 
         Caption         =   "Modificar Producto"
      End
      Begin VB.Menu menu_01_03 
         Caption         =   "Eliminar Producto"
      End
      Begin VB.Menu menu_01_01 
         Caption         =   "Seleccionar Productos"
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu menu2_1 
         Caption         =   "Agregar Producto"
      End
      Begin VB.Menu menu2_3 
         Caption         =   "Modificar Producto"
      End
      Begin VB.Menu menu2_2 
         Caption         =   "Eliminar Producto"
      End
   End
   Begin VB.Menu menu3 
      Caption         =   "Menu3"
      Visible         =   0   'False
      Begin VB.Menu menu3_1 
         Caption         =   "Ver Productos"
      End
   End
End
Attribute VB_Name = "FrmCronoProduccion2_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim xNomMatPriPro As String
Dim QueHace As Integer
Dim Agregando As Boolean
Dim RstLis As New ADODB.Recordset
Dim RstMatPro As New ADODB.Recordset
Dim xIdMatPri As Integer
Dim xFchPro, xHorPro As Date

Dim oPDF As cPDF
Dim xNumPag As Integer
Dim xFilaInicial As Integer
Dim xHorIni As Date                     ' ALMACENARA LA HORA DE INICIO CUANDO SE CREA O MODIFICA UN REGISTRO
Dim IdMenuActivo As Integer             ' INDICA EL CODIGO DEL MENU ACTIVO
Dim fOrdenLista As Boolean              ' especfica el orden de la lista de la consulta
Dim SeEjecuto As Boolean

Dim visEvent As Boolean
Dim modifEvent As Boolean
Dim agregEvent As Boolean

Dim mIdRegistro& 'identificador del registro
Dim OrigFX As Long
Dim OrigFY As Long

Dim DETECTOR_ As CalendarHitTestInfo
Dim EVENTO_ As CalendarEvent

' Variables para las Propiedades de Procesado
Dim MODO_TAREA As Integer           ' 0 = "Al finalizar"; 1 = "Al porcentaje"; 2 = "Al minuto"; 3 = "Linea"
Dim PORCENTAJE As Double
Dim MINUTOS_ As String
Dim INCLUIR_HORAS As Integer        ' 0 = "Incluir"; 1 = "No incluir"
Dim HOR_INI As String
Dim HOR_FIN As String

Dim CORR_ As Double

Dim RstProductos As New ADODB.Recordset
Dim RstProductosAux As New ADODB.Recordset

Dim RstPersonal As New ADODB.Recordset
Dim RstPersonalAux As New ADODB.Recordset

Dim RstTareas As New ADODB.Recordset
Dim RstTareasAux As New ADODB.Recordset

Dim IDCR_ As Double
Dim cSQL As String
Dim CAMBIO_ As Boolean
Dim ARRASTRANDO_ As Boolean
Dim CARGO_ As Boolean
Dim VERIFICO_ As Boolean
Dim con_SQLS As ADODB.Connection        ' Conexion Base de datos del control de asistencia


'*****************************************************************************************************
'* Descripcion      : EVITA LA EDICION DEL CALENDARIO EN DIVERSAS SITUACIONES
'* Modificacion     : 15/02/11 JOSE CHACON
'*****************************************************************************************************
Private Sub CalendarControl1_BeforeEditOperation(ByVal OpParams As XtremeCalendarControl.CalendarEditOperationParameters, bCancelOperation As Boolean)
    If QueHace = 3 Then bCancelOperation = True: Exit Sub
    ' SI ES EDITAR EL CONTENIDO MANUAL SE CANCELA
    If OpParams.Operation = xtpCalendarEO_EditSubject_ByMouseClick Then bCancelOperation = True
    ' sI SE EDITA POR LA TECLA F2 SE CANCELA
    If OpParams.Operation = xtpCalendarEO_EditSubject_ByF2 Then bCancelOperation = True
    ' SI SE CAMBIA DE TAMAÑO MANUALMENTE EL EVENTO SE CANCELA
    If OpParams.Operation = xtpCalendarEO_DragResizeBegin Then bCancelOperation = True
    If OpParams.Operation = xtpCalendarEO_DragResizeEnd Then bCancelOperation = True
    ' EDITAR DESPUES DE UN CAMBIO DE TAMAÑO SE CANCELA
    If OpParams.Operation = xtpCalendarEO_EditSubject_AfterEventResize Then bCancelOperation = True
    ' Editar despues de un arrastre
    If OpParams.Operation = xtpCalendarEO_DragMove Then ARRASTRANDO_ = True Else ARRASTRANDO_ = False
    ' Eliminacion Manual
    If OpParams.Operation = xtpCalendarEO_DeleteEvent Then bCancelOperation = True
End Sub

Private Sub CalendarControl1_DblClick()
    visEvent = True
    If TxtTipPro.Text = 1 Then
        menu_01_01_Click
        Set DETECTOR_ = Nothing
    Else
        mostrarFormulario False, True, False
        Set DETECTOR_ = Nothing
    End If
End Sub

Private Sub CalendarControl1_KeyDown(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    
On Error Resume Next
    'Se activa el detector para la vista activa del calendario
    Set DETECTOR_ = CalendarControl1.ActiveView.HitTest
    'Se agrega el evento del detector
    Set EVENTO_ = DETECTOR_.ViewEvent.Event
    
    If KeyCode = vbKeyInsert Then
        Menu2_1_Click
    End If
    If KeyCode = vbKeyDelete Then
        'Si el detector no tiene evento activo
        If DETECTOR_.ViewEvent Is Nothing Then Exit Sub
        menu2_2_Click
    End If
    If KeyCode = 113 Then
        'Si el detector no tiene evento activo
        If DETECTOR_.ViewEvent Is Nothing Then Exit Sub
        Menu2_3_Click
    End If
    Set DETECTOR_ = Nothing
End Sub

Private Sub CalendarControl1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim xRs As New ADODB.Recordset
    
On Error GoTo ERROR_
            
    If ARRASTRANDO_ Then
        Dim EVENTO_AUX As CalendarEvent
        Dim EVENTOINICIAL_ As CalendarEvent
        Dim CANTIDAD_ As Double
        Dim IDLINEA_ As Double
        Dim IDCRDET_ As Double
        Dim IDITEM_ As Double
        Dim HORINI_ As String
        Dim FECHINI_ As Date
           
        Set EVENTOINICIAL_ = EVENTO_
           
        Set DETECTOR_ = CalendarControl1.ActiveView.HitTest
        If DETECTOR_.ViewEvent Is Nothing Then
            RstProductos.Filter = adFilterNone
            ' Se limpia el calendario
            CalendarControl1.DataProvider.RemoveAllEvents
            ' Se llenan todos los eventos sin modificar
            LlenarDatos
            Exit Sub
        End If
        Set EVENTO_ = DETECTOR_.ViewEvent.Event
        
        
        If EVENTOINICIAL_ <> EVENTO_ Then
            RstProductos.Filter = adFilterNone
            ' Se limpia el calendario
            CalendarControl1.DataProvider.RemoveAllEvents
            ' Se llenan todos los eventos sin modificar
            LlenarDatos
            Exit Sub
        End If
    
        ' Se llena el evento auxiliar
        Set EVENTO_AUX = EVENTO_
        ' Se elimina el evento arrastrado
        CalendarControl1.DataProvider.DeleteEvent EVENTO_
        
        ' Se determina que evento se esta trabajando
        IDCRDET_ = NulosN(EVENTO_AUX.ReminderSoundFile)
                                
        ' Se filtra el producto relacionado
        If RstProductos.State = 0 Then Exit Sub
        RstProductos.Filter = "id = " & IDCRDET_ & ""
        If RstProductos.RecordCount = 0 Then Exit Sub
        If RstProductosAux.State = 0 Then DEFINIR_RST_TMP RstProductosAux, RstProductos
        limpiarRST RstProductosAux, False
        CARGAR_RST_TMP RstProductosAux, RstProductos
        
        ' Se filtran las tareas relacionadas
        If RstTareas.State = 0 Then Exit Sub
        RstTareas.Filter = "idcrdet = " & IDCRDET_ & ""
        If RstTareas.RecordCount = 0 Then Exit Sub
        If RstTareasAux.State = 0 Then DEFINIR_RST_TMP RstTareasAux, RstTareas
        limpiarRST RstTareasAux, False
        CARGAR_RST_TMP RstTareasAux, RstTareas
        
        ' Se modifica los datos del producto
        ' se determina la nueva hora y fecha de inicio
        RstProductosAux("fchpro") = Format(EVENTO_AUX.StartTime, "dd/mm/yyyy")
        RstProductosAux("horpro") = Format(EVENTO_AUX.StartTime, "HH:mm")
        
        ' Se calculan los valores del evento
        IDLINEA_ = NulosN(RstTareasAux("idlinea"))
        IDCRDET_ = NulosN(RstTareasAux("idcrdet"))
        IDITEM_ = NulosN(RstProductosAux("iditem"))
        CANTIDAD_ = calcularRdmto(IDLINEA_, IDCRDET_, RstTareasAux, NulosN(RstProductosAux("cantidad")))
        HORINI_ = Format(RstProductosAux("horpro"), "HH:mm")
        FECHINI_ = CDate(RstProductosAux("fchpro"))
        ' Se carga el recordset auxiliar
        DEFINIR_RST_TMP xRs, RstTareasAux
        CARGAR_RST_TMP xRs, RstTareasAux
                
        procesarCronograma xRs, False, CANTIDAD_, HORINI_, HORINI_, FECHINI_, IDITEM_, IDCRDET_, IDLINEA_
        
        RstTareasAux.Filter = "idcrdet = " & IDCRDET_ & " And activo = True"
        RstTareasAux.MoveLast
        
        ' se determina la nueva hora y fecha de fin
        RstProductosAux("fchfin") = RstTareasAux("fchfin")
        RstProductosAux("horfin") = RstTareasAux("horfintar")
        
        Dim Rpta As Integer
        'xTitulo = "Desplazar Evento"
        Rpta = MsgBox("¿Se moverá el evento a esta nueva ubicación; desea continuar?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
        If Rpta = vbYes Then
            ' Para Tareas
            RstTareas.Filter = "idcrdet = " & IDCRDET_
            RstTareasAux.Filter = "idcrdet = " & IDCRDET_
            limpiarRST RstTareas, False
            CARGAR_RST_TMP RstTareas, RstTareasAux
            ' Los productos
            RstProductos.Filter = "id = " & IDCRDET_
            RstProductosAux.Filter = "id = " & IDCRDET_
            limpiarRST RstProductos, False
            CARGAR_RST_TMP RstProductos, RstProductosAux
            RstProductos.Filter = adFilterNone
            
            ' Se limpia el calendario
            CalendarControl1.DataProvider.RemoveAllEvents
            ' Se llenan todos los eventos
            LlenarDatos
        Else
            RstProductos.Filter = adFilterNone
            ' Se limpia el calendario
            CalendarControl1.DataProvider.RemoveAllEvents
            ' Se llenan todos los eventos sin modificar
            LlenarDatos
        End If
    End If
    ARRASTRANDO_ = False
    Exit Sub
ERROR_:
    'xTitulo = "Error al Desplazar"
    MsgBox "Ocurrio un error al desplazar el evento; intente de nuevo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Sub

Private Sub CalendarControl1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Set DETECTOR_ = CalendarControl1.ActiveView.HitTest
    Set EVENTO_ = DETECTOR_.ViewEvent.Event
    
    If Button = 2 Then
        If QueHace <> 3 Then
            If NulosN(TxtTipPro.Text) = 1 Then
                PopupMenu menu_01
            Else
                PopupMenu menu2
            End If
        Else
            If NulosN(TxtTipPro.Text) = 1 Then
                If DETECTOR_.ViewEvent Is Nothing Then Exit Sub
                PopupMenu menu3
            End If
        End If
    End If
End Sub

Private Sub CalendarControl1_SelectionChanged(ByVal SelType As XtremeCalendarControl.CalendarSelectionChanged)
    If SelType = xtpCalendarSelectionDays Then
        If Not CARGO_ Then Exit Sub ' Si no ha cargado el calendario
        If VERIFICO_ Then Exit Sub ' Si se verfico que corresponde al rango de fechas
        
        Dim FCHINI_ As Date
        Dim FCHFIN_ As Date
        Dim TODODIA_ As Boolean
        Dim PRIMERDIASEMANA_ As Date
        Dim ULTIMODIASEMANA_ As Date

        ' Se obtienen los datos del dia seleccionado
        CalendarControl1.ActiveView.GetSelection FCHINI_, FCHFIN_, TODODIA_
        
        ' SI es una fecha Incoherente
        If Format(FCHINI_, "yyyy") < AnoTra Then Exit Sub
        
        PRIMERDIASEMANA_ = CDate(TxtFchIni.valor)
        ULTIMODIASEMANA_ = CDate(TxtFchFin.valor)
        FCHINI_ = Format(FCHINI_, "dd/mm/yyyy")

        If FCHINI_ < PRIMERDIASEMANA_ Or FCHINI_ > ULTIMODIASEMANA_ Then
            CalendarControl1.ActiveView.ShowDay (PRIMERDIASEMANA_)
            VERIFICO_ = True
            'OptVista_Click 0
            CalendarControl1.ViewType = xtpCalendarFullWeekView
        End If
        VERIFICO_ = False
    End If
End Sub

Sub CrearCabeceraVS(numPag As Integer, Optional PROGRAMADOR_ As String)
    Dim xCad As String

    FrmVsPrinter.Vs.TextAlign = taLeftTop
    FrmVsPrinter.Vs.FontName = "Courier New"
    FrmVsPrinter.Vs.FontBold = True
    FrmVsPrinter.Vs.FontSize = 9
    
    FrmVsPrinter.Vs.CurrentX = 900:      FrmVsPrinter.Vs.CurrentY = 600
    FrmVsPrinter.Vs.Paragraph = "PROGRAMADOR   : " & NulosC(PROGRAMADOR_)

    FrmVsPrinter.Vs.CurrentX = 7950:      FrmVsPrinter.Vs.CurrentY = 600
    FrmVsPrinter.Vs.Paragraph = "FECHA        : " & Format(Date, "dd/mm/yy")

    FrmVsPrinter.Vs.CurrentX = 7950:      FrmVsPrinter.Vs.CurrentY = 800
    FrmVsPrinter.Vs.Paragraph = "Nº Pagina    : " & Format(numPag, "0000")

    FrmVsPrinter.Vs.DrawLine 900, 1050, 11000, 1050
End Sub

Private Function hallarCaracTareas(IDLINEA_ As Double, IDTAREA_ As Double, _
                                        Optional UNIDXHOR_ As Boolean = True, _
                                        Optional EFICIENCIA_ As Boolean = False) As String
    Dim xRs As New ADODB.Recordset
    Dim mensaje As String
    Dim campo As String
    
    If UNIDXHOR_ Then campo = "kghora"
    If EFICIENCIA_ Then campo = "efictar"
    
    cSQL = "SELECT pro_lineadet.idlinea, pro_lineadet.idtar, pro_lineadet.kghora, pro_lineadet.efictar, pro_lineadet.numopideal, pro_lineadet.durtarreal, pro_lineadet.eficop " _
        + vbCr + "From pro_lineadet " _
        + vbCr + "GROUP BY pro_lineadet.idlinea, pro_lineadet.idtar, pro_lineadet.kghora, pro_lineadet.efictar, pro_lineadet.numopideal, pro_lineadet.durtarreal, pro_lineadet.eficop " _
        + vbCr + "HAVING (((pro_lineadet.idlinea)=" & IDLINEA_ & ") AND ((pro_lineadet.idtar)=" & IDTAREA_ & "));"
    
    RST_Busq xRs, cSQL, xCon

    If xRs.State = 0 Then mensaje = ""
    If xRs.RecordCount = 0 Then
        mensaje = ""
    Else
        mensaje = xRs("" & campo & "")
    End If
    
    hallarCaracTareas = mensaje
End Function

Private Sub ImprimirLinea(RstProd As ADODB.Recordset, RstTar As ADODB.Recordset, _
                                RstPers As ADODB.Recordset)
    Dim A As Integer
    Dim B As Integer
    Dim xLinea As Integer           ' Fila de impresion
    Dim xColumna As Integer         ' Columna de impresion
    Dim numPag As Integer           ' Numero de pagina de Impresion
    Dim numper As Double            ' Numero de Personas
    Dim ID_LINEA As Double
    Dim Rst As New ADODB.Recordset
    
    With FrmVsPrinter.Vs
        numPag = 0
        .BrushColor = &H80000005
        .FontSize = 11
        .TextAlign = taCenterMiddle
                
        RstProd.MoveFirst
        On Error Resume Next
        For A = 1 To RstProd.RecordCount
            xLinea = 1300
            xColumna = 900
            numPag = numPag + 1
            If A > 1 Then .NewPage
            CrearCabeceraVS numPag, LblSupervisor.Caption
                        
            '******************************************************************* Titulo
            .FontSize = 12
            .FontBold = True
            .TextAlign = taCenterMiddle
            
            .TextBox "LINEA DE PRODUCCION", xColumna, xLinea, 8000, 500, True, False, True
            .FontSize = 10
            .TextAlign = taCenterTop
            .TextBox "NUM. PROD.", xColumna + 8100, xLinea, 1900, 250, True, False, True
            xLinea = xLinea + 240
            .TextBox NulosC(RstProd("numprod")), xColumna + 8100, xLinea, 1900, 250, True, False, True
            
            .TextAlign = taLeftMiddle
            .FontSize = 9
            
            '******************************************************************* Detalle de la Linea
            xLinea = xLinea + 300
            .FontBold = True
            .TextBox "Detalles de la Solicitud", xColumna, xLinea, 3500, 250, True, False, False
            
'            '*************************************************************************
            .FontBold = False
            xLinea = xLinea + 250
            .TextBox "Producto", xColumna, xLinea, 1500, 250, True, False, False
            .TextBox NulosC(RstProd("descripcion")), xColumna + 1500, xLinea, 7000, 250, True, False, False
            
            '*************************************************************************
            xLinea = xLinea + 250
            .TextBox "Fecha Prog.", xColumna, xLinea, 1500, 250, True, False, False
            .TextBox Format(RstProd("fchpro"), "dd/mm/yyyy"), xColumna + 1500, xLinea, 6000, 250, True, False, False
            
            RstTar.Filter = "idcrdet = " & NulosN(RstProd("id")) & " And activo = True"
            RstTar.MoveFirst
            
            '*************************************************************************
            .TextBox "Cantidad", xColumna + 7500, xLinea, 1000, 250, True, False, False
            .TextBox Format(RstProd("cantidad"), "0.00") & " " & encontrarUnidad(RstProd("iditem")), xColumna + 8550, xLinea, 6000, 250, True, False, False
            
            '*************************************************************************
            xLinea = xLinea + 250
            .TextBox "Responsable ", xColumna, xLinea, 1500, 250, True, False, False
            .TextBox RstTar("nomresp"), xColumna + 1500, xLinea, 6000, 250, True, False, False
            
            Dim xFila As Integer
             '******************************************************************* Detalle de la Linea
            xLinea = xLinea + 300
            .TextAlign = taLeftMiddle
            .FontBold = True
            .TextBox "Detalles de la Linea", xColumna, xLinea, 2500, 250, True, False, False
            '*************************************************************************
            
            .FontBold = False
            xLinea = xLinea + 350
            .TextAlign = taCenterMiddle
            .TextBox "Ord.", xColumna, xLinea, 500, 500, True, False, True
            .TextBox "Tarea", xColumna + 500, xLinea, 3500, 500, True, False, True
            .TextBox "Durac.", xColumna + 4000, xLinea, 800, 500, True, False, True
            .TextBox "Hor.Ini", xColumna + 4800, xLinea, 800, 500, True, False, True
            .TextBox "Hor.Fin", xColumna + 5600, xLinea, 800, 500, True, False, True
            .TextBox "Num. Pers.", xColumna + 6400, xLinea, 800, 500, True, False, True
            .TextBox "Unid.x Hora", xColumna + 7200, xLinea, 1000, 500, True, False, True
            .TextBox "%Rdto", xColumna + 8200, xLinea, 800, 500, True, False, True
            .TextBox "Cant. Proc.", xColumna + 9000, xLinea, 1000, 500, True, False, True
            
            numper = 0
            xLinea = xLinea + 500
            xFila = xLinea
            
            ID_LINEA = NulosN(RstTar("idlinea"))
            
            For B = 1 To RstTar.RecordCount
                .FontSize = 8
                .FontBold = False
                
                .TextAlign = taLeftMiddle
                .TextBox " " & NulosN(RstTar("idtar")), xColumna, xLinea, 500, 250, True, False, True
                .TextBox " " & NulosC(RstTar("destar")), xColumna + 500, xLinea, 3500, 250, True, False, True
                .TextAlign = taCenterMiddle
                .TextBox Format(RstTar("durtar"), "HH:mm"), xColumna + 4000, xLinea, 800, 250, True, False, True
                .TextBox Format(RstTar("horinitar"), "HH:mm"), xColumna + 4800, xLinea, 800, 250, True, False, True
                .TextBox Format(RstTar("horfintar"), "HH:mm"), xColumna + 5600, xLinea, 800, 250, True, False, True
                .TextBox Format(RstTar("numper"), "00"), xColumna + 6400, xLinea, 800, 250, True, False, True
                
                .TextAlign = taRightMiddle
                .TextBox Format(hallarCaracTareas(NulosN(RstTar("idlinea")), RstTar("idtar")), "0.00") & " ", xColumna + 7200, xLinea, 1000, 250, True, False, True
                .TextBox Format(RstTar("aplpor"), "0.00") & "% ", xColumna + 8200, xLinea, 800, 250, True, False, True
                .TextBox Format(RstTar("cantproc"), "0.00") & " ", xColumna + 9000, xLinea, 1000, 250, True, False, True
                
                numper = numper + NulosN(RstTar("numper"))
                
                RstTar.MoveNext
                If RstTar.EOF = True Then Exit For
                
                xLinea = xLinea + 250
                
                If xLinea >= 16200 Then
                    xLinea = 1300
                    numPag = numPag + 1
                    .NewPage
                    CrearCabeceraVS numPag
                End If
            Next B
            
            xLinea = xLinea + 250
            .TextAlign = taRightMiddle
            .TextBox "TOTAL", xColumna, xLinea, 4000, 250, True, False, True
            .TextAlign = taCenterMiddle
            .TextBox Format(numper, "00"), xColumna + 6400, xLinea, 800, 250, True, False, True
            
            .FontBold = False
            xLinea = xLinea + 400
            .TextBox "CANTIDAD", xColumna + 4750, xLinea, 1500, 250, True, False, True
            
            .TextAlign = taCenterMiddle
            xLinea = xLinea + 250
            .FontSize = 7
            .TextBox calcularProdAnterior(ID_LINEA, True, True), xColumna + 500, xLinea, 4250, 250, True, False, True
            .FontSize = 8
            .TextBox "", xColumna + 4750, xLinea, 1500, 250, True, False, True
            
            .TextAlign = taLeftMiddle
            .TextBox " Hora Ini.", xColumna + 6500, xLinea, 1500, 250, True, False, True
            .TextBox "", xColumna + 7500, xLinea, 1500, 250, True, False, True
            '*************************************************************************
            
            .TextAlign = taCenterMiddle
            xLinea = xLinea + 250
            .TextBox "P1", xColumna, xLinea, 500, 250, True, False, True
            .FontSize = 7
            .TextBox NulosC(RstProd("descripcion")), xColumna + 500, xLinea, 4250, 250, True, False, True
            .FontSize = 8
            .TextBox "", xColumna + 4750, xLinea, 1500, 250, True, False, True
            
            .TextAlign = taLeftMiddle
            .TextBox " Hora Fin", xColumna + 6500, xLinea, 1500, 250, True, False, True
            .TextBox "", xColumna + 7500, xLinea, 1500, 250, True, False, True
            '*************************************************************************
            xLinea = xLinea + 250
            .TextAlign = taCenterMiddle
            .TextBox "P2", xColumna, xLinea, 500, 250, True, False, True
            .TextBox "", xColumna + 500, xLinea, 4250, 250, True, False, True
            .TextBox "", xColumna + 4750, xLinea, 1500, 250, True, False, True
            '*************************************************************************
            xLinea = xLinea + 250
            .TextBox "P3", xColumna, xLinea, 500, 250, True, False, True
            .TextBox "", xColumna + 500, xLinea, 4250, 250, True, False, True
            .TextBox "", xColumna + 4750, xLinea, 1500, 250, True, False, True
            
            .TextAlign = taLeftMiddle
            .TextBox " Lote", xColumna + 6500, xLinea, 1500, 250, True, False, True
            .TextBox "", xColumna + 7500, xLinea, 1500, 250, True, False, True
            '*************************************************************************
            
            '****************************************************************************************
            '******************************************************************* Detalle del Personal
            '****************************************************************************************
            xLinea = xLinea + 300
            .TextAlign = taLeftMiddle
            .FontBold = True
            .TextBox "Detalles del Personal", xColumna, xLinea, 2500, 250, True, False, False
            '*************************************************************************
                        
            xLinea = xLinea + 350
            
            .FontBold = False
            .TextAlign = taCenterMiddle
            .TextBox "Item", xColumna, xLinea, 500, 500, True, False, True
            .TextBox "PERSONAL", xColumna + 500, xLinea, 3500, 500, True, False, True
            .TextBox "Tarea", xColumna + 4000, xLinea, 800, 500, True, False, True
            .TextBox "Hr.Ini.", xColumna + 4800, xLinea, 1000, 500, True, False, True
            .TextBox "Hr.Ter.", xColumna + 5800, xLinea, 1000, 500, True, False, True
            .TextBox "M.P.", xColumna + 6800, xLinea, 800, 500, True, False, True
            .TextBox "Prod01", xColumna + 7600, xLinea, 800, 500, True, False, True
            .TextBox "Prod02", xColumna + 8400, xLinea, 800, 500, True, False, True
            .TextBox "Prod03", xColumna + 9200, xLinea, 800, 500, True, False, True
            
            RstPers.Filter = "idcrdet = " & NulosN(RstProd("id")) & ""
            If RstPers.RecordCount <> 0 Then RstPers.MoveFirst
        
            xLinea = xLinea + 500
            xFila = xLinea
            
            ' Se agrega 5 campos mas para ingresar personal
            numper = numper + 5
            ' Se verifica que se muestre no menos de 25 personas
            ' If numper < 25 Then numper = 25
            For B = 1 To numper
                .FontSize = 10
                .FontBold = False
                .TextAlign = taLeftMiddle
                
                .TextBox " " & Format(B, "00"), xColumna, xLinea, 500, 300, True, False, True
                If Not RstPers.EOF Then
                    .TextBox " " & NulosC(RstPers("nombre")), xColumna + 500, xLinea, 3500, 300, True, False, True
                    .TextBox " " & NulosC(RstPers("idtar")), xColumna + 4000, xLinea, 800, 300, True, False, True
                    .TextBox "", xColumna + 4800, xLinea, 1000, 300, True, False, True
                    .TextBox "", xColumna + 5800, xLinea, 1000, 300, True, False, True
                    .TextBox "", xColumna + 6800, xLinea, 800, 300, True, False, True
                    .TextBox "", xColumna + 7600, xLinea, 800, 300, True, False, True
                    .TextBox "", xColumna + 8400, xLinea, 800, 300, True, False, True
                    .TextBox "", xColumna + 9200, xLinea, 800, 300, True, False, True
                    RstPers.MoveNext
                Else
                    .TextBox "", xColumna + 500, xLinea, 3500, 300, True, False, True
                    .TextBox "", xColumna + 4000, xLinea, 800, 300, True, False, True
                    .TextBox "", xColumna + 4800, xLinea, 1000, 300, True, False, True
                    .TextBox "", xColumna + 5800, xLinea, 1000, 300, True, False, True
                    .TextBox "", xColumna + 6800, xLinea, 800, 300, True, False, True
                    .TextBox "", xColumna + 7600, xLinea, 800, 300, True, False, True
                    .TextBox "", xColumna + 8400, xLinea, 800, 300, True, False, True
                    .TextBox "", xColumna + 9200, xLinea, 800, 300, True, False, True
                End If
                
                xLinea = xLinea + 300
                
                If xLinea >= 16200 Then
                    xLinea = 1300
                    numPag = numPag + 1
                    .NewPage
                    CrearCabeceraVS numPag
                End If
            Next B
            
            '****************************************************************************************
            '************************************************************************** Observaciones
            '****************************************************************************************
            xLinea = xLinea + 100
            
            If xLinea >= 15500 Then
                xLinea = 1300
                numPag = numPag + 1
                .NewPage
                CrearCabeceraVS numPag
            End If
            
            .TextAlign = taLeftMiddle
            .FontBold = True
            .TextBox "Observaciones", xColumna, xLinea, 2500, 250, True, False, False
            '*************************************************************************
            xLinea = xLinea + 450
            .DrawLine xColumna + 500, xLinea, 10000, xLinea
            xLinea = xLinea + 250
            .DrawLine xColumna + 500, xLinea, 10000, xLinea
            xLinea = xLinea + 250
            .DrawLine xColumna + 500, xLinea, 10000, xLinea
            xLinea = xLinea + 250
            .DrawLine xColumna + 500, xLinea, 10000, xLinea
            
SIGUIENTE:
        Next A
    End With
End Sub


Private Sub imprimir(TIPO_ As Integer)
    'TIPO_ = 0:LINEA
    'TIPO_ = 1:ACABADO
    'TIPO_ = 2:REPORTE
    Dim xLinea As Integer
    Dim xform As New eps_librerias.FormSeleccion
    Dim xRs As New ADODB.Recordset
    Dim nSQLFiltro As String '--Almacenara el filtro por movimiento
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(6, 5) As String
    
    Select Case TIPO_
        Case 0
            xCampos(0, 0) = "Fch. Prog.":    xCampos(0, 1) = "fchpro":         xCampos(0, 2) = "950":     xCampos(0, 3) = "D":    xCampos(0, 4) = "D"
            xCampos(1, 0) = "Producto":      xCampos(1, 1) = "descripcion":    xCampos(1, 2) = "3200":    xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
            xCampos(2, 0) = "Supervisor":    xCampos(2, 1) = "nombre":         xCampos(2, 2) = "2800":    xCampos(2, 3) = "C":    xCampos(2, 4) = "C"
            xCampos(3, 0) = "Cantidad":      xCampos(3, 1) = "cantidad":       xCampos(3, 2) = "900":     xCampos(3, 3) = "N":    xCampos(3, 4) = "N"
            xCampos(4, 0) = "Hr. Ini.":      xCampos(4, 1) = "horpro":         xCampos(4, 2) = "1100":    xCampos(4, 3) = "C":    xCampos(4, 4) = "C"
            xCampos(5, 0) = "Hr. Fin":       xCampos(5, 1) = "horfin":         xCampos(5, 2) = "1100":    xCampos(5, 3) = "C":    xCampos(5, 4) = "C"

            'consulta para obtener listado de Productos
            cSQL = "SELECT 0 AS xsel, pro_cronogramadet.fchpro, alm_inventario.descripcion, pro_cronogramadet.cantidad, pro_cronogramadet.horpro, pro_cronogramadet.horfin, pro_cronogramadet.id, pro_cronogramadet.idcr, pro_cronograma.semana, pla_empleados.nombre " _
                + vbCr + "FROM ((pro_cronograma RIGHT JOIN (pro_cronogramadet LEFT JOIN alm_inventario ON pro_cronogramadet.iditem = alm_inventario.id) ON pro_cronograma.id = pro_cronogramadet.idcr) LEFT JOIN pro_cronogramatarea ON pro_cronogramadet.id = pro_cronogramatarea.idcrdet) LEFT JOIN pla_empleados ON pro_cronogramatarea.idresp = pla_empleados.id " _
                + vbCr + "GROUP BY 0, pro_cronogramadet.fchpro, alm_inventario.descripcion, pro_cronogramadet.cantidad, pro_cronogramadet.horpro, pro_cronogramadet.horfin, pro_cronogramadet.id, pro_cronogramadet.idcr, pro_cronograma.semana, pla_empleados.nombre " _
                + vbCr + "Having (((pro_cronograma.semana) = " & NulosN(ComboSemanas.Text) & ")) " _
                + vbCr + "ORDER BY pro_cronogramadet.fchpro;"
            
            xform.SQLCad = cSQL
                
            xform.titulo = "Operaciones a Imprimir"
            Set xform.Coneccion = xCon
            Set xRs = Nothing
            Set xRs = xform.seleccionar(xCampos)
            
            If xRs.State = 0 Then Exit Sub
            If xRs.RecordCount = 0 Then Exit Sub
            
            xRs.MoveFirst
            With FrmVsPrinter.Vs
                .StartDoc
                Me.MousePointer = vbHourglass
                Dim A As Integer
                For A = 1 To xRs.RecordCount
                    If A > 1 Then .NewPage
                    RstProductos.Filter = adFilterNone
                    RstProductos.Filter = "id = " & NulosN(xRs("id"))
                    If RstProductos.RecordCount = 0 Then GoTo SIGUIENTE
                    ImprimirLinea RstProductos, RstTareas, RstPersonal
SIGUIENTE:
                xRs.MoveNext
                Next A
                Me.MousePointer = vbDefault
                .EndDoc
            End With
            
        Case 1
    End Select
    'Muestra la preimagen de la impresion
    FrmVsPrinter.WindowState = 2
    FrmVsPrinter.Show
End Sub

Private Sub iniciarCampos()
    Dim A As Integer
    Dim pTema2007 As CalendarThemeOffice2007
    'Se guarda el tema del calendario activo
    Set pTema2007 = CalendarControl1.Theme
    
    'Se cambia el color de seleccion
    pTema2007.WeekView.Day.BackgroundSelectedColor = &HFF0000
    
    pTema2007.Event.Normal.Location.Color = &HFF&
    
    pTema2007.Event.Normal.Body.Color = &HFF0000
    pTema2007.Event.Normal.Body.Font.Size = 8
    
    pTema2007.Event.Selected.Background.ColorDark = &HFF0000
    pTema2007.Event.Selected.BorderColor = &HFF&
    
    pTema2007.Event.Selected.Subject.Color = &HFF&
    pTema2007.Event.Selected.Subject.Font.Size = 7
    pTema2007.Event.Normal.Subject.Font.Size = 7
    pTema2007.Event.Normal.Subject.Font.Bold = True
    
    ' Se habilita los mensajes de ayuda
    CalendarControl1.EnableToolTips True
    ' Se deshabilita el ingreso de eventos por mouse
    CalendarControl1.Options.EnableAddNewTooltip = False
    
    
    CalendarControl1.DayView.TimeScale = 60
    CalendarControl1.DayView.MinColumnWidth = 408
    CalendarControl1.DayView.EnableHScroll False
    
    ARRASTRANDO_ = False
    CARGO_ = False
    VERIFICO_ = False
    
    TabOne1.CurrTab = 0
    
    'se cargan las semanas
    For A = 1 To 52
        ComboSemanas.AddItem A
    Next A
    
    MODO_TAREA = 3 ' Procesar segun Linea
    PORCENTAJE = 10
    MINUTOS_ = "00:10"
    INCLUIR_HORAS = False ' No incluir Horas de refrigerio
    HOR_INI = "13:00"
    HOR_FIN = "14:00"
    CORR_ = -666
    
    fg(0).Editable = flexEDKbdMouse
    fg(0).ColWidth(8) = 0
    fg(0).ColWidth(9) = 0
    fg(0).ColWidth(10) = 0
    fg(0).ColWidth(11) = 0
    fg(0).ColWidth(12) = 0
    
    fg(1).ColWidth(4) = 0
    fg(1).ColWidth(5) = 0
    fg(1).ColWidth(6) = 0
    fg(1).ColWidth(7) = 0
    
    fg(2).ColWidth(0) = 0
    fg(2).ColWidth(5) = 0
    fg(2).ColWidth(6) = 0
    fg(2).ColWidth(7) = 0
End Sub

Private Sub Cmd_Click(Index As Integer)
    Dim xFrm As New sgi2_produccion.produccion
    
    Select Case Index
        Case 0 ' Agregar Producto
            agregarCampos True, False
            
        Case 1 ' Establecer propiedades de procesado
            aplicarPropiedades False, True
            centrarFrm Frame9
            Frame9.ZOrder 0
            Frame9.Visible = True
            
        Case 2 ' Procesar la linea
            Dim xRs As New ADODB.Recordset
            
            RstTareasAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) & ""
            
            If RstTareasAux.RecordCount = 0 Then
                cSQL = "SELECT " & NulosN(LblIdCrDet.Caption) & " AS idcrdet, pro_receta.iditem, pro_lineadet.idtar, pro_lineadet.orden, pro_tareas.descripcion AS destar, " & NulosN(TxtCant.Text) & " AS cantidad, pro_lineadet.factor, pro_lineadet.kghora AS costokg, pro_lineadet.numop AS numper, pro_lineadet.intervalo AS horarr, pro_lineadet.rdmto AS aplpor, '" & Format(DTPHoras.Value, "HH:mm") & "' AS horinitar, '" & Format(LblDia.Caption, "dd/mm/yyyy") & "' AS fchini, -1 AS activo " _
                    + vbCr + "FROM (pro_lineadet LEFT JOIN pro_tareas ON pro_lineadet.idtar = pro_tareas.id) LEFT JOIN (pro_receta LEFT JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) ON pro_lineadet.idrec = pro_receta.id " _
                    + vbCr + "Where (((pro_lineadet.idlineadet) = " & NulosN(TxtIdLineaDet.Text) & ")) " _
                    + vbCr + "ORDER BY pro_lineadet.orden;"
                    
                RST_Busq xRs, cSQL, xCon
                
                If xRs.State = 0 Then Exit Sub
                
                If xRs.RecordCount = 0 Then
                    'xTitulo = "Error al Procesar Tareas"
                    MsgBox "No se encontro datos de la Linea de Produccion; Agregue una y procese de nuevo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    Exit Sub
                End If
            
                procesarCronograma xRs, , , , , , , NulosN(LblIdCrDet.Caption), NulosN(TxtIdLineaDet.Text)
            Else
                Dim CANTIDAD_ As Double
                Dim IDLINEA_ As Double
                Dim IDCRDET_ As Double
                Dim IDITEM_ As Double
                Dim HORINI_ As String
                Dim FECHINI_ As Date
                
                DEFINIR_RST_TMP xRs, RstTareasAux
                CARGAR_RST_TMP xRs, RstTareasAux
                
                IDLINEA_ = NulosN(TxtIdLineaDet.Text)
                IDCRDET_ = NulosN(LblIdCrDet.Caption)
                IDITEM_ = NulosN(TxtMatProd.Text)
                CANTIDAD_ = calcularRdmto(IDLINEA_, IDCRDET_, RstTareasAux, NulosN(TxtCant.Text))
                HORINI_ = Format(DTPHoras.Value, "HH:mm")
                FECHINI_ = CDate(Format(LblDia.Caption, "dd/mm/yyyy"))
                
                procesarCronograma xRs, False, CANTIDAD_, HORINI_, HORINI_, FECHINI_, IDITEM_, IDCRDET_, IDLINEA_
            End If
            ' Se actualizan las Tareas
            pCargarDatos fg(0), False, True, , , False
        
        Case 3 ' Agregar tarea
            agregarCampos False, True
            ' Se carga al personal relacionado con esa tarea si es que lo hubiera
            RstPersonalAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) & " And idtar = " & NulosN(LblIdTarea) & ""
            pCargarDatos fg(1), True, False, , , False
            
        Case 4 ' Agregar Personal
            procesarPersonal True, False, False, False
            
        Case 5 ' Listar personal
            procesarPersonal False, True, False, False
            
        Case 6 ' Eliminar Personal
            procesarPersonal False, False, True, False
            
        Case 7 ' Eliminar Todos
            procesarPersonal False, False, False, True
            
        Case 8 ' Ver Ranking
            LbNumSel.Caption = 0
            OptSel(1).Value = True
            ' Se procesa el ranking para mostrarlo
            procesarRanking
            
        Case 9 ' Editar Linea de Produccion
            ' Se llama al formulario de linea de produccion
            xFrm.CronogramaMantLinea xCon
            Set xFrm = Nothing
        
        Case 10 ' Acepta Agregar/Modificar Detalle
            If NulosC(TxtNumProd.Text) = "" Then
                'xTitulo = "Numero de Produccion"
                MsgBox "Ingrese un Numero de Produccion, para la programacion actual", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            ' Se actualiza el estado como estado actual
            ' Para Tareas
            RstTareas.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption)
            RstTareasAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption)
            limpiarRST RstTareas, False
            CARGAR_RST_TMP RstTareas, RstTareasAux
            ' Para personal
            RstPersonal.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption)
            RstPersonalAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption)
            limpiarRST RstPersonal, False
            CARGAR_RST_TMP RstPersonal, RstPersonalAux
            ' Los productos
            ' Se agrega o se modifica en el registro de Productos
            Dim FECHAINI_ As Date
            Dim FECHAFIN_ As Date
            
            FECHAINI_ = CDate(Format(LblDia.Caption, "dd/mm/yyyy") & " " + Format(DTPHoras.Value, "HH:mm"))
            FECHAFIN_ = CDate(Format(TxtBFchFin.valor, "dd/mm/yyyy") & " " + Format(DTPHoraFin.Value, "HH:mm"))
            
            If RstProductosAux.State = 0 Then DEFINIR_RST_TMP RstProductosAux, RstProductos
            limpiarRST RstProductosAux
            RstProductosAux.AddNew
            RstProductosAux("id") = NulosN(LblIdCrDet.Caption)
            RstProductosAux("idcr") = NulosN(LblIdCrDet.Caption)
            RstProductosAux("numprod") = NulosC(TxtNumProd.Text)
            RstProductosAux("fchpro") = FECHAINI_
            RstProductosAux("fchfin") = FECHAFIN_
            RstProductosAux("horpro") = Format(FECHAINI_, "HH:mm")
            RstProductosAux("horfin") = Format(FECHAFIN_, "HH:mm")
            RstProductosAux("iditem") = NulosN(TxtMatProd.Text)
            '*****************************************************************
            RstProductosAux("idrec") = NulosN(lblIdRec.Caption)
            RstProductosAux("codrec") = NulosC(TxtCodRec.Text)
            '*****************************************************************
            RstProductosAux("cantidad") = NulosN(TxtCant.Text)
            RstProductosAux("descripcion") = LblMatProd.Caption
            RstProductosAux("abrev") = LblUnidad.Caption
            RstProductosAux.Update
            
            RstProductos.Filter = "id = " & NulosN(LblIdCrDet.Caption)
            RstProductosAux.Filter = "id = " & NulosN(LblIdCrDet.Caption)
            limpiarRST RstProductos, False
            CARGAR_RST_TMP RstProductos, RstProductosAux
            
            ' Se Agrega en el calendario
            operaciones
            
        Case 11 ' Cancela Agregar/Modificar Detalle
            ' Se retorna a un estado anterior
            ' Para Tareas
            RstTareas.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption)
            RstTareasAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption)
            limpiarRST RstTareasAux, False
            CARGAR_RST_TMP RstTareasAux, RstTareas
            ' Para Personal
            RstPersonal.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption)
            RstPersonalAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption)
            limpiarRST RstPersonalAux, False
            CARGAR_RST_TMP RstPersonalAux, RstPersonal
            
            visEvent = False
            FrmAdd.Visible = False
            
        Case 12 ' Aceptar Propiedades de procesado
            aplicarPropiedades True
            Frame9.Visible = False
            
        Case 13 ' Cancela Propiedades de procesado
            Frame9.Visible = False
            
        Case 14 ' Adicionar de Ranking
            procesarRanking False, True
            
        Case 15 ' Cancela Procesar Ranking
            Frame10.Visible = False
        
        Case 16 ' Elegir Receta
            agregarCampos False, False, False, True
            'cargarReceta NulosN(LbIdRec.Caption), NulosN(TxtCant.Text)
            
        Case 17 ' Mostrar Formulario Receta de Produccion
            xFrm.MamRecetas xCon
            
        Case 18 ' Escoger Encargado de Linea
            agregarCampos False, False, True
        
        Case 19 ' Imprimir Linea
            With FrmVsPrinter.Vs
                .StartDoc
                Me.MousePointer = vbHourglass
            
                RstProductos.Filter = adFilterNone
                RstProductos.Filter = "id = " & NulosN(LblIdCrDet.Caption)
                If RstProductos.RecordCount = 0 Then Exit Sub
                ImprimirLinea RstProductos, RstTareas, RstPersonal
                
                Me.MousePointer = vbDefault
                .EndDoc
            End With
            'Muestra la preimagen de la impresion
            FrmVsPrinter.WindowState = 2
            FrmVsPrinter.Show
            
        Case 20 ' Buscar Linea
            agregarCampos False, False, False, False, True
            
    End Select
End Sub

Private Sub conectar_BD_SQL_Server(nombre_BD As String)
    Dim AP_PROVIDER As String
    Dim AP_INITIALCATALOG As String
    Dim AP_DATASOURCE As String
    Dim AP_USER As String
    Dim AP_PASSWORD As String

    ' La conexión a la base de datos
    Set con_SQLS = New ADODB.Connection
    
    AP_PROVIDER = LeerLineaINI(Trim(App.Path) + "\asistencia.ini", "PROVIDER", "DATOS")
    AP_INITIALCATALOG = LeerLineaINI(Trim(App.Path) + "\asistencia.ini", "INITIALCATALOG", "DATOS")
    AP_DATASOURCE = LeerLineaINI(Trim(App.Path) + "\asistencia.ini", "DATASOURCE", "DATOS")
    AP_USER = LeerLineaINI(Trim(App.Path) + "\asistencia.ini", "USER", "DATOS")
    AP_PASSWORD = LeerLineaINI(Trim(App.Path) + "\asistencia.ini", "PASSWORD", "DATOS")
    
    con_SQLS.Open "Provider=" & AP_PROVIDER & "; " & _
             "Initial Catalog=" & AP_INITIALCATALOG & "; " & _
             "Data Source=" & AP_DATASOURCE & "; " & _
             "user id = " & AP_USER & "; " & _
             "password = " & AP_PASSWORD & ""
End Sub

Private Sub procesarRanking(Optional MOSTRAR_ As Boolean = True, Optional AGREGAR_ As Boolean = False)
    Dim RstRanking As New ADODB.Recordset
    Dim A As Integer
    Dim nSQLId_0 As String
    Dim nSQLId_1 As String
    Dim FECHA_ As Date
    Dim REINTENTO_ As Boolean
    
    If MOSTRAR_ Then
On Error GoTo ERROR_AL_MOSTRAR
        
        LblProd2.Caption = LblMatProd.Caption
        LblTarea2.Caption = LblTarea.Caption
        
        conectar_BD_SQL_Server "TEMPUS"
        
        ' Generar la lista de personal para no considerar en la lista
        RstPersonalAux.Filter = "idcrdet = " & LblIdCrDet.Caption & ""
        nSQLId_0 = GENERAR_SQL_ID_RST(RstPersonalAux, "idper", " AND pro_controltardet.idref", "NOT IN", True)
        FECHA_ = Date
        REINTENTO_ = False
REINTENTAR:
        nSQLId_1 = GENERAR_SQL_ID_RST(buscarAsistencia(FECHA_), "DNI", " AND pla_empleados.numdoc", "IN", False)
        
        If nSQLId_1 = "" Then
            If Not REINTENTO_ Then
                REINTENTO_ = True
                FECHA_ = Date - 1
                GoTo REINTENTAR
            End If
            'xTitulo = "Error al Procesar Asistencia"
            MsgBox "No se encontro datos de la Asistencia; Se mostrara a todo el Personal", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
        
'        cSQL = "SELECT pro_controltardet.tipo, pro_controltardet.idref AS idper, pla_empleados.nombre, pro_receta.iditem, alm_inventario.descripcion AS producto, pro_controltardet.idtar, pro_tareas.abrev AS tarea, Sum(pro_controltardet.cant) AS SumaDecant, Last(pro_controltar.fchtra) AS ÚltimoDefchtra, Sum(1) AS diasTrab, pla_empleados.numdoc " _
'            + vbCr + "FROM pro_controltar INNER JOIN (pro_tareas RIGHT JOIN (mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN ((pro_receta RIGHT JOIN pro_controltardet ON pro_receta.id = pro_controltardet.idrec) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON alm_inventario.id = pro_receta.iditem) ON mae_unidades.id = pro_controltardet.idunimed) ON pro_tareas.id = pro_controltardet.idtar) ON pro_controltar.id = pro_controltardet.idctr " _
'            + vbCr + "GROUP BY pro_controltardet.tipo, pro_controltardet.idref, pla_empleados.nombre, pro_receta.iditem, alm_inventario.descripcion, pro_controltardet.idtar, pro_tareas.abrev, alm_inventario.descripcion, pro_tareas.abrev, pla_empleados.numdoc " _
'            + vbCr + "Having (((pro_controltardet.Tipo) = 1) And ((pla_empleados.nombre) Is Not Null) And ((pro_receta.iditem) = " & NulosN(TxtMatProd.Text) & ") And ((pro_controltardet.idtar) = " & NulosN(LblIdTarea.Caption) & ") And ((pro_tareas.abrev) Is Not Null)) " & nSQLId_0 & nSQLId_1 _
'            + vbCr + "ORDER BY alm_inventario.descripcion, pro_tareas.abrev; " _
'            + vbCr + "Union " _
'            + vbCr + "SELECT pro_controltardet.tipo, pro_controltardetgr.idper, pla_empleados.nombre, pro_receta.iditem, alm_inventario.descripcion AS producto, pro_controltardet.idtar, pro_tareas.abrev AS tarea, Sum(pro_controltardetgr.cant) AS SumaDecant, Last(pro_controltar.fchtra) AS ÚltimoDefchtra, Sum(1) AS diasTrab, pla_empleados.numdoc " _
'            + vbCr + "FROM pro_controltar INNER JOIN (((pro_tareas RIGHT JOIN (mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN (pro_receta RIGHT JOIN pro_controltardet ON pro_receta.id = pro_controltardet.idrec) ON alm_inventario.id = pro_receta.iditem) ON mae_unidades.id = pro_controltardet.idunimed) ON pro_tareas.id = pro_controltardet.idtar) LEFT JOIN pro_controltardetgr ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id) ON pro_controltar.id = pro_controltardet.idctr " _
'            + vbCr + "GROUP BY pro_controltardet.tipo, pro_controltardetgr.idper, pla_empleados.nombre, pro_receta.iditem, alm_inventario.descripcion, pro_controltardet.idtar, pro_tareas.abrev, alm_inventario.descripcion, pro_tareas.abrev, pla_empleados.numdoc " _
'            + vbCr + "HAVING (((pro_controltardet.tipo)=2) AND ((pla_empleados.nombre) Is Not Null) AND ((pro_receta.iditem)= " & NulosN(TxtMatProd.Text) & ") AND ((pro_controltardet.idtar)= " & NulosN(LblIdTarea.Caption) & ") AND ((pro_tareas.abrev) Is Not Null)) " & nSQLId_0 & nSQLId_1 _

        cSQL = "SELECT pro_controltardet.tipo, pro_controltardet.idref AS idper, pla_empleados.nombre, pro_receta.iditem, alm_inventario.descripcion AS producto, pro_controltardet.idtar, pro_tareas.abrev AS tarea, Sum(pro_controltardet.cant) AS SumaDecant, Last(pro_controltar.fchtra) AS ÚltimoDefchtra, Sum(1) AS diasTrab, pla_empleados.numdoc " _
            + vbCr + "FROM pro_controltar INNER JOIN (pro_tareas RIGHT JOIN (mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN ((pro_receta RIGHT JOIN pro_controltardet ON pro_receta.id = pro_controltardet.idrec) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON alm_inventario.id = pro_receta.iditem) ON mae_unidades.id = pro_controltardet.idunimed) ON pro_tareas.id = pro_controltardet.idtar) ON pro_controltar.id = pro_controltardet.idctr " _
            + vbCr + "GROUP BY pro_controltardet.tipo, pro_controltardet.idref, pla_empleados.nombre, pro_receta.iditem, alm_inventario.descripcion, pro_controltardet.idtar, pro_tareas.abrev, alm_inventario.descripcion, pro_tareas.abrev, pla_empleados.numdoc " _
            + vbCr + "Having (((pro_controltardet.Tipo) = 1) And ((pla_empleados.nombre) Is Not Null) And ((pro_receta.iditem) = " & NulosN(TxtMatProd.Text) & ") And ((pro_controltardet.idtar) = " & NulosN(LblIdTarea.Caption) & ") And ((pro_tareas.abrev) Is Not Null)) " & nSQLId_0 & nSQLId_1 _
            + vbCr + "ORDER BY alm_inventario.descripcion, pro_tareas.abrev;"
        
        RST_Busq RstRanking, cSQL, xCon
        RstRanking.Sort = "SumaDecant Desc"
        
        fg(2).Rows = 1
        If RstRanking.State = 0 Then Exit Sub
        
        If RstRanking.RecordCount <> 0 Then
            ' Se llenan los Datos
            RstRanking.MoveFirst
            For A = 1 To RstRanking.RecordCount
                fg(2).Rows = fg(2).Rows + 1
                fg(2).TextMatrix(A, 1) = 0
                fg(2).TextMatrix(A, 2) = A
                fg(2).TextMatrix(A, 3) = RstRanking("numdoc")
                fg(2).TextMatrix(A, 4) = RstRanking("nombre")
                fg(2).TextMatrix(A, 5) = RstRanking("iditem")
                fg(2).TextMatrix(A, 6) = RstRanking("idtar")
                fg(2).TextMatrix(A, 7) = RstRanking("idper")
                fg(2).TextMatrix(A, 8) = Format(RstRanking("SumaDecant"), "0.00")
                fg(2).TextMatrix(A, 9) = RstRanking("diasTrab")
                fg(2).TextMatrix(A, 10) = Format(NulosN(fg(2).TextMatrix(A, 8) / fg(2).TextMatrix(A, 9)), "0.00")
                RstRanking.MoveNext
                If RstRanking.EOF Then Exit For
            Next A
            
            ' Se ordena segun eficiencia
            fg(2).Select 1, 10
            fg(2).Sort = flexSortNumericDescending
            For A = 1 To fg(2).Rows - 1
                fg(2).TextMatrix(A, 2) = A
            Next A
        End If
        centrarFrm Frame10
        ' Se pone en primer plano
        Frame10.ZOrder 0
        Frame10.Visible = True
        Exit Sub
ERROR_AL_MOSTRAR:
        MsgBox "Ocurrio un Error al Visualizar, verifique que el Servidor este activo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    If AGREGAR_ Then
        Dim contador As Integer
        Dim num As Double
        
        num = NulosN(LblNTrab.Caption) - (fg(1).Rows - 1)
        If num <= 0 Then
            MsgBox "La Tarea actual no puede admitir mas Personal", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        
        For A = 1 To fg(2).Rows - 1
            If num <= 0 Then
                MsgBox "La Tarea actual no puede admitir mas Personal, solo se agregara al personal necesario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit For
            End If
            If fg(2).TextMatrix(A, 1) = 0 Then GoTo SIGUIENTE
            ' agregando los datos al rst temporal
            RstPersonalAux.AddNew
            RstPersonalAux("idcrdet") = NulosN(LblIdCrDet.Caption)
            RstPersonalAux("idtar") = NulosN(LblIdTarea.Caption)
            RstPersonalAux("idper") = fg(2).TextMatrix(A, 7)
            RstPersonalAux("codigo") = ""
            RstPersonalAux("nombre") = fg(2).TextMatrix(A, 4)
            RstPersonalAux("activo") = True
            num = num - 1
SIGUIENTE:
        Next A
        RstPersonalAux.Filter = adFilterNone
        pCargarDatos fg(1), , , , , False
        Frame10.Visible = False
    End If
End Sub

Private Function buscarAsistencia(FECHA_CONSULTA As Date) As ADODB.Recordset
    ' El recordset para acceder a los datos
    Dim RstAsistencia As ADODB.Recordset
    
    ' Datos para la consulta
    Dim CONS_FECH_ASISTENCIA As String
    
    ' CONSULTA DE FECHA DE ASISTENCIA
    CONS_FECH_ASISTENCIA = "(TEMPUS.MARCACIONES.FECHA = CAST('" & FECHA_CONSULTA & "' AS datetime)) "
    
    Set RstAsistencia = New ADODB.Recordset
    
    ' CONSULTA
    cSQL = "SELECT TEMPUS.EMPRESAS.NOMBRE AS EMP, " _
                    + vbCr + "TEMPUS.PERSONAL.APELLIDO_PATERNO + ' ' + TEMPUS.PERSONAL.APELLIDO_MATERNO + ' ' + TEMPUS.PERSONAL.NOMBRES AS NOMPER, " _
                    + vbCr + "CONVERT(varchar(12), TEMPUS.PERSONAL.FECHA_DE_NACIMIENTO, 103) AS FECHNAC, CONVERT(varchar(12), " _
                    + vbCr + "TEMPUS.PERSONAL.FECHA_DE_INGRESO, 103) AS FECHING, TEMPUS.PERSONAL.DNI, CONVERT(varchar(12), TEMPUS.MARCACIONES.FECHA, 103) AS FECHMARC, " _
                    + vbCr + "CONVERT(varchar(10), TEMPUS.MARCACIONES.HORA, 108) AS HORMARC, TEMPUS.CARGOS.DESCRIPCION " _
            + vbCr + "FROM TEMPUS.MARCACIONES INNER JOIN " _
                    + vbCr + "TEMPUS.PERSONAL ON TEMPUS.MARCACIONES.CODIGO = TEMPUS.PERSONAL.CODIGO AND " _
                    + vbCr + "TEMPUS.MARCACIONES.EMPRESA = TEMPUS.PERSONAL.EMPRESA INNER JOIN " _
                    + vbCr + "TEMPUS.EMPRESAS ON TEMPUS.PERSONAL.EMPRESA = TEMPUS.EMPRESAS.EMPRESA INNER JOIN " _
                    + vbCr + "TEMPUS.CARGOS ON TEMPUS.PERSONAL.CARGO = TEMPUS.CARGOS.CARGO " _
            + vbCr + "WHERE " & CONS_FECH_ASISTENCIA & " " _
            + vbCr + "ORDER BY TEMPUS.MARCACIONES.FECHA, TEMPUS.PERSONAL.APELLIDO_PATERNO"
    
    ' Abrir el recordset de forma estática, no vamos a cambiar datos
    RST_Busq RstAsistencia, cSQL, con_SQLS
    
    Set buscarAsistencia = RstAsistencia
End Function

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

Private Sub procesarPersonal(AGREGAR_ As Boolean, LISTAR_ As Boolean, _
                                    ELIMINAR_ As Boolean, ELIMINARTODOS_ As Boolean)
    If QueHace = 3 Then Exit Sub
    
    Dim nSQL As String
    Dim nSQLId As String
    Dim nSQLTmp  As String
    Dim nTitulo As String
    Dim xform As New eps_librerias.FormSeleccion
    Dim xRs As New ADODB.Recordset
    Dim xCampos() As String
    Dim A As Integer
        
    If AGREGAR_ Then
        ReDim xCampos(3, 4) As String
        
        xCampos(0, 0) = "Cod. Empleado":        xCampos(0, 1) = "codemp":      xCampos(0, 2) = "2000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
        xCampos(1, 0) = "Apellidos y Nombres":  xCampos(1, 1) = "nombre":      xCampos(1, 2) = "4000":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
        xCampos(2, 0) = "Area":                 xCampos(2, 1) = "area":        xCampos(2, 2) = "2000":     xCampos(2, 3) = "C":    xCampos(2, 4) = "C"
        
        If fg(1).Rows - 1 = NulosN(LblNTrab.Caption) Then MsgBox "La Tarea actual no puede admitir mas Personal", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo: Exit Sub
        If fg(1).Rows = fg(1).FixedRows Then fg(1).Rows = fg(1).Rows + 1
    
        ' generar la lista de personal para no considerar en la lista
        RstPersonalAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) & ""
        nSQLId = GENERAR_SQL_ID_RST(RstPersonalAux, "idper", " AND pla_empleados.id", "NOT IN", True)
        ' generar la consulta
        nSQL = "SELECT pla_empleados.codigo AS codemp, pla_empleados.id AS idemp, pla_empleados.nombre, -1 AS activo, mae_area.descripcion AS area " _
            + vbCr + "FROM ((pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN mae_area ON pla_empleados.idarea = mae_area.id " _
            + vbCr + "Where (((pro_empdet.idfun) = 6)) " & nSQLId _
            + vbCr + "ORDER BY pla_empleados.nombre;"
            
        nTitulo = "Buscando Personal"
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
            
        xform.titulo = "Buscando Personal"
        
        If xRs.State = 0 Then Exit Sub
        
        ' agregando los datos al rst temporal
        RstPersonalAux.AddNew
        RstPersonalAux("idcrdet") = NulosN(LblIdCrDet.Caption)
        RstPersonalAux("idtar") = NulosN(LblIdTarea.Caption)
        RstPersonalAux("destar") = NulosC(LblTarea.Caption)
        RstPersonalAux("activo") = xRs("activo")
        RstPersonalAux("idper") = xRs("idemp")
        RstPersonalAux("nombre") = xRs("nombre")
        RstPersonalAux("codigo") = xRs("codemp")
        
        RstPersonalAux.Update
        pCargarDatos fg(1), True, False, , , False
        
        Agregando = False
        Set xform = Nothing
        Set xRs = Nothing
    End If
    
    If LISTAR_ Then
        Dim num As Integer ' numero de registros que se van a agregar
        ReDim xCampos(3, 4) As String
        
        num = NulosN(LblNTrab.Caption) - (fg(1).Rows - 1)
        If num <= 0 Then
            MsgBox "La Tarea actual no puede admitir mas Personal", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        
        xCampos(0, 0) = "Cod. Empleado":        xCampos(0, 1) = "codemp":      xCampos(0, 2) = "2000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
        xCampos(1, 0) = "Apellidos y Nombres":  xCampos(1, 1) = "nombre":      xCampos(1, 2) = "5000":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
        xCampos(2, 0) = "Area":                 xCampos(2, 1) = "area":        xCampos(2, 2) = "2000":     xCampos(2, 3) = "C":    xCampos(2, 4) = "C"
        
        If fg(1).Rows = fg(1).FixedRows Then fg(1).Rows = fg(1).Rows + 1
    
        ' generar la lista de personal para no considerar en la lista
        RstPersonalAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) & ""
        nSQLId = GENERAR_SQL_ID_RST(RstPersonalAux, "idper", " AND pla_empleados.id", "NOT IN", True)
        ' generar la consulta
        nSQL = "SELECT 0 AS xsel, pla_empleados.codigo AS codemp, pla_empleados.id AS idemp, pla_empleados.nombre, -1 AS activo, mae_area.descripcion AS area " _
            + vbCr + "FROM ((pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN mae_area ON pla_empleados.idarea = mae_area.id " _
            + vbCr + "Where (((pro_empdet.idfun) = 6)) " & nSQLId _
            + vbCr + "ORDER BY pla_empleados.nombre;"
            
        nTitulo = "Buscando Personal"
    
        xform.SQLCad = nSQL
            
        xform.titulo = "Buscando Personal"
        Set xform.Coneccion = xCon
        Set xRs = Nothing
        Set xRs = xform.seleccionar(xCampos)
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        For A = 1 To num
            ' agregando los datos al rst temporal
            RstPersonalAux.AddNew
            RstPersonalAux("idcrdet") = NulosN(LblIdCrDet.Caption)
            RstPersonalAux("idtar") = NulosN(LblIdTarea.Caption)
            RstPersonalAux("destar") = NulosC(LblTarea.Caption)
            RstPersonalAux("activo") = xRs("activo")
            RstPersonalAux("idper") = xRs("idemp")
            RstPersonalAux("nombre") = xRs("nombre")
            RstPersonalAux("codigo") = xRs("codemp")
            
            xRs.MoveNext
            If xRs.EOF = True Then Exit For
        Next A
        RstPersonalAux.Filter = adFilterNone
        pCargarDatos fg(1), True, False, , , False
        
        Set xform = Nothing
        Set xRs = Nothing
    End If
    
    If ELIMINAR_ Then
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
        
        If Agregando Then
            If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
        End If
        
        If RstPersonalAux.RecordCount <> 0 Then RstPersonalAux.MoveFirst
        
        Do While Not RstPersonalAux.EOF
            If RstPersonalAux.RecordCount = 0 Then Exit Do
            If NulosN(RstPersonalAux("idper")) = NulosN(fg(1).TextMatrix(fg(1).Row, 5)) Then
                RstPersonalAux.Delete
                Exit Do
            End If
            RstPersonalAux.MoveNext
        Loop
        
        pCargarDatos fg(1), True, False, , , False
    End If
    
    If ELIMINARTODOS_ Then
        num = fg(1).Rows - 1
        For A = 1 To num
            Agregando = False
            fg(1).Select 1, 1, 1, fg(1).Cols - 1
            procesarPersonal False, False, True, False
        Next A
        pCargarDatos fg(1), True, False, , , False
    End If
End Sub

Private Sub aplicarPropiedades(MODIFICAR_ As Boolean, Optional CARGAR_ As Boolean = False)
    If MODIFICAR_ Then
        If OptTarea(0).Value = True Then MODO_TAREA = 0
        If OptTarea(1).Value = True Then MODO_TAREA = 1
        If OptTarea(2).Value = True Then MODO_TAREA = 2
        If OptTarea(3).Value = True Then MODO_TAREA = 3
        
        If OptHoras(0).Value = True Then INCLUIR_HORAS = 0
        If OptHoras(1).Value = True Then INCLUIR_HORAS = 1
        
        PORCENTAJE = NulosN(TxtPctje.Text)
        MINUTOS_ = Format(DTPMinutos.Value, "HH:mm")
        HOR_INI = Format(DTPHorIni.Value, "HH:mm")
        HOR_FIN = Format(DTPHorFin.Value, "HH:mm")
    End If
    
    If CARGAR_ Then
        OptTarea(MODO_TAREA).Value = True
        OptHoras(INCLUIR_HORAS).Value = True
        TxtPctje.Text = PORCENTAJE
        DTPMinutos.Value = MINUTOS_
        DTPHorIni.Value = HOR_INI
        DTPHorFin.Value = HOR_FIN
    End If
End Sub

Private Sub CmdAcepta_Click(Index As Integer)
    If Index = 0 Then
        Dim B As Integer
        
        If NulosN(TxtTotal.Text) > NulosN(TxtCan.Text) Then
            MsgBox "El cantidad a procesar en productos es mayor a la cantidad de materia prima", vbInformation + vbOKOnly + vbDefaultButton1
            TxtTotal.SetFocus
            Exit Sub
        End If
        
        For B = 1 To Fg2.Rows - 1
            RstMatPro.Filter = adFilterNone
            If Abs(NulosN(Fg2.TextMatrix(B, 3))) = 1 Then
                RstMatPro.Filter = "iditem = " & xIdMatPri & " AND fchpro = " & xFchPro & " AND horpro = " & Format(xHorPro, "HH:mm") & " AND idpro = " & NulosN(Fg2.TextMatrix(B, 4)) & ""
                If RstMatPro.RecordCount = 0 Then
                    RstMatPro.AddNew
                    RstMatPro("idcr") = 0
                    RstMatPro("iditem") = xIdMatPri
                    RstMatPro("fchpro") = xFchPro
                    RstMatPro("horpro") = xHorPro
                    RstMatPro("idpro") = Fg2.TextMatrix(B, 4)
                    RstMatPro("cantidad") = NulosN(Fg2.TextMatrix(B, 2))
                Else
                    RstMatPro("cantidad") = NulosN(Fg2.TextMatrix(B, 2))
                End If
            Else
                If NulosN(xIdMatPri) = 0 Then Exit Sub
                RstMatPro.Filter = "iditem = " & xIdMatPri & " AND fchpro = " & xFchPro & " AND horpro = " & Format(xHorPro, "HH:mm") & " AND idpro = " & NulosN(Fg2.TextMatrix(B, 4)) & ""
                If RstMatPro.RecordCount <> 0 Then
                    RstMatPro.Delete
                End If
            End If
        Next B
        
        CmdCancelar_Click 0
    End If
End Sub

Private Sub cargarReceta(ID_RECETA As Integer, CANTIDAD As Double)
    Dim A As Integer
    Dim xRs As New ADODB.Recordset
    
    cSQL = "SELECT pro_recetains.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro, [pro_recetains]![canpro]* " & CANTIDAD & " AS canreq " _
        + vbCr + " FROM (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
        + vbCr + " WHERE (((pro_recetains.idrec)=" & ID_RECETA & "));"
    
    RST_Busq xRs, cSQL, xCon

    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    If RstRecetaAux.State = 0 Then Exit Sub
    
    For A = 1 To xRs.RecordCount
        RstRecetaAux.AddNew
        
        RstRecetaAux("idcrdet") = NulosN(LblIdCrDet.Caption)                    ' idcrdet
        RstRecetaAux("fchprog") = NulosC(LblDia.Caption)                        ' fecha de programacion
        RstRecetaAux("idpro") = NulosN(TxtMatProd.Text)                         ' ID Producto
        RstRecetaAux("idrec") = NulosN(LbIdRec.Caption)                         ' ID Receta
        RstRecetaAux("nomrec") = NulosC(LblReceta.Caption)                      ' Descripcion Receta
        RstRecetaAux("codrec") = NulosC(TxtCodRec.Text)                         ' Codigo Receta
        RstRecetaAux("canpro") = NulosN(TxtCant.Text)                           ' Cantidad de Producto
        RstRecetaAux("idresp") = NulosN(TxtIdEncarg.Text)                       ' id responsable
        RstRecetaAux("nomresp") = NulosC(LblEncargado.Caption)                  ' nombre responsable
        RstRecetaAux("idins") = NulosN(xRs("iditem"))                           ' ID Insumo
        RstRecetaAux("canins") = NulosN(xRs("canreq"))                          ' Cantidad Insumo
        RstRecetaAux("nomins") = NulosC(xRs("descripcion"))                     ' Descripcion Insumo
        RstRecetaAux("abrev") = NulosC(xRs("abrev"))                            ' Abrev de UM
        
        RstRecetaAux.Update
        
        xRs.MoveNext
        If xRs.EOF = True Then Exit For
    Next A
    pCargarDatos fg(3), False, False, False, True
End Sub

Private Sub agregarCampos(PRODUCTO_ As Boolean, TAREA_ As Boolean, _
                        Optional RESPONSABLE_ As Boolean = False, _
                        Optional RECETA_ As Boolean = False, _
                        Optional LINEA_ As Boolean = False)
    Dim xCampos() As String
    Dim RstLinea As New ADODB.Recordset
    
    If PRODUCTO_ Then
        ReDim xCampos(2, 4) As String
        Dim xRs As New ADODB.Recordset
        Dim titulo As String
        Dim Rpta As Integer
        
        If QueHace = 3 Then Exit Sub
    
        'descripcion                     'campo                       'tamaño                         'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "despro":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Uni. Med.":     xCampos(1, 1) = "abrev":     xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"

        cSQL = "SELECT pro_receta.iditem, pro_receta.id AS idrec, pro_receta.codrec, alm_inventario.descripcion AS despro, mae_unidades.abrev " _
            + vbCr + "FROM (pro_receta LEFT JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_receta.idunimed = mae_unidades.id " _
            + vbCr + "WHERE (((pro_receta.prirec)=1));"
            
        titulo = "Buscando Productos"
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos, titulo, "despro", "despro"
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        RstTareasAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) & ""
        If RstTareasAux.RecordCount > 0 Then
            'xTitulo = "Seleccion de Producto"
            Rpta = MsgBox("¿Se Eliminara Todo el Personal y Tareas Relacionado a la linea Anterior; desea continuar?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbNo Then Exit Sub
        End If
        
        ' Se Limpia las Tareas Relacionadas con el evento
        RstTareasAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) & ""
        limpiarRST RstTareasAux, False
        
        ' Se Limpia el Personal Relacionado con el evento
        RstPersonalAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) & ""
        limpiarRST RstPersonalAux, False
        
        ' Se llena el detalle
        ' Producto
        TxtMatProd.Text = NulosN(xRs("iditem"))
        LblMatProd.Caption = NulosC(xRs("despro"))
        LblUnidad.Caption = NulosC(xRs("abrev"))
        ' Receta
        TxtCodRec.Text = NulosC(xRs("codrec"))
        lblIdRec.Caption = NulosN(xRs("idrec"))
        
        TxtCant.SetFocus

        ' Se verifica si el producto seleccionado tiene una linea activa
        cSQL = "SELECT pro_linea.id AS idlineadet, pro_linea.descripcion " _
                + vbCr + "From pro_linea " _
                + vbCr + "WHERE (((pro_linea.idrec)=" & NulosN(xRs("idrec")) & ") AND ((pro_linea.activo)=-1));"
                        
        RST_Busq RstLinea, cSQL, xCon
        
        If RstLinea.State = 0 Then GoTo ERROR_AL_ENCONTRAR_LINEA
        If RstLinea.RecordCount = 0 Then GoTo ERROR_AL_ENCONTRAR_LINEA
        
        ' Se llena la linea Activa
        TxtIdLineaDet.Text = NulosN(RstLinea("idlineadet"))
        LblLinea.Caption = NulosC(RstLinea("descripcion"))
        
        Set xRs = Nothing
        Set RstLinea = Nothing
        Exit Sub
        
ERROR_AL_ENCONTRAR_LINEA:
        MsgBox "El producto procesado no tiene Linea activa, procese una para calcular tiempos de Producción", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        
        TxtIdLineaDet.Text = ""
        LblLinea.Caption = ""
        Set xRs = Nothing
        Set RstLinea = Nothing
    End If
    
    If TAREA_ Then
        ReDim xCampos(2, 4) As String
        Dim nTitulo As String
        Dim nSQLId As String
    
        'descripcion                    'campo                            'tamaño                         'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Tarea":        xCampos(0, 1) = "destar":         xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Id Tarea":     xCampos(1, 1) = "idtar":          xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
                   
        ' Se filtra las tareas no seleccionadas
        RstTareasAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) & " And activo = False"

        nSQLId = GENERAR_SQL_ID_RST(RstTareasAux, "idtar", " AND pro_tareas.id", "NOT IN", True)
        
        cSQL = "SELECT pro_tareas.id AS idtar, pro_lineadet.orden, pro_tareas.descripcion AS destar, pro_lineadet.numop " _
            + vbCr + "FROM pro_lineadet LEFT JOIN pro_tareas ON pro_lineadet.idtar = pro_tareas.id " _
            + vbCr + "Where (((pro_lineadet.idlineadet) = " & NulosN(TxtIdLineaDet.Text) & ")) " & nSQLId _
            + vbCr + "GROUP BY pro_tareas.id, pro_lineadet.orden, pro_tareas.descripcion, pro_lineadet.numop " _
            + vbCr + "ORDER BY pro_lineadet.orden;"
    
        nTitulo = "Buscando Tareas"
        Dim RstTmp As New ADODB.Recordset
        CARGAR_DLL_EPSBUSCAR xCon, RstTmp, cSQL, xCampos(), nTitulo, "orden", "descripcion", Principio
    
        If RstTmp.State = 0 Then Exit Sub
        If RstTmp.RecordCount = 0 Then Exit Sub
        
        TxtOrden.Text = NulosN(RstTmp("idtar"))
        LblTarea.Caption = NulosC(RstTmp("destar"))
        LblIdTarea.Caption = NulosN(RstTmp("idtar"))
        LblNTrab.Caption = NulosN(RstTmp("numop"))
        LblDetTrab.Caption = fg(1).Rows - 1 & " de " & NulosN(LblNTrab.Caption)
        
        pCargarDatos fg(1), True, False, , , False
        
        Set RstTmp = Nothing
    End If
    
    If RESPONSABLE_ Then
        ReDim xCampos(2, 4) As String
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Apellidos y Nombres":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":                xCampos(1, 1) = "id":        xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
                
        cSQL = "SELECT pro_emp.idemp, pro_emp.sup, pro_emp.prog, pro_emp.res, pla_empleados.nombre " _
            + vbCr + "FROM (pro_emp LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id " _
            + vbCr + "Where (((pro_empdet.idfun) = 3)) " _
            + vbCr + "GROUP BY pro_emp.idemp, pro_emp.sup, pro_emp.prog, pro_emp.res, pla_empleados.nombre " _
            + vbCr + "Having (((pla_empleados.nombre) Is Not Null)) " _
            + vbCr + "ORDER BY pla_empleados.nombre;"
            
        nTitulo = "Buscando Personal Encargado"
                
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        LblEncargado.Caption = NulosC(xRs("nombre"))     ' codigo de la receta
        TxtIdEncarg.Text = NulosN(xRs("idemp"))          ' ID de la receta
        Cmd(20).SetFocus
        
        ' Se llena el detalle en las Tareas
        RstTareasAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) & ""
        If RstTareasAux.RecordCount = 0 Then Exit Sub
        RstTareasAux.MoveFirst
        While Not RstTareasAux.EOF
            RstTareasAux("idresp") = NulosN(TxtIdEncarg.Text)
            RstTareasAux("nomresp") = NulosC(LblEncargado.Caption)
            RstTareasAux.MoveNext
        Wend
    End If
    
    If RECETA_ Then ' Cargar Receta
        ReDim xCampos(2, 4) As String
        
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Receta":     xCampos(1, 1) = "codrec":        xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
        
        cSQL = "SELECT pro_receta.codrec, pro_receta.descripcion, pro_receta.prirec, pro_receta.id " _
            + vbCr + "From pro_receta " _
            + vbCr + "Where (((pro_receta.iditem) = " & NulosN(TxtMatProd.Text) & ")) " _
            + vbCr + "ORDER BY pro_receta.prirec;"
            
        nTitulo = "Buscando Recetas del Producto"
                
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        TxtCodRec.Text = NulosC(xRs("codrec"))             ' Codigo de la receta
        lblIdRec.Caption = NulosN(xRs("id"))               ' Id de la receta
        
        
        ' Se verifica si el producto seleccionado tiene una linea activa
        cSQL = "SELECT pro_linea.id AS idlineadet, pro_linea.descripcion " _
                + vbCr + "From pro_linea " _
                + vbCr + "WHERE (((pro_linea.idrec)=" & NulosN(xRs("id")) & ") AND ((pro_linea.activo)=-1));"
                        
        RST_Busq RstLinea, cSQL, xCon
        
        If RstLinea.State = 0 Then GoTo ERROR_AL_ENCONTRAR_LINEA2
        If RstLinea.RecordCount = 0 Then GoTo ERROR_AL_ENCONTRAR_LINEA2
        
        ' Se llena la linea Activa
        TxtIdLineaDet.Text = NulosN(RstLinea("idlineadet"))
        LblLinea.Caption = NulosC(RstLinea("descripcion"))
        
        Set xRs = Nothing
        Set RstLinea = Nothing
        Exit Sub
        
ERROR_AL_ENCONTRAR_LINEA2:
        MsgBox "El producto procesado no tiene Linea activa, procese una para calcular tiempos de Producción", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        
        TxtIdLineaDet.Text = ""
        LblLinea.Caption = ""
        Set xRs = Nothing
        Set RstLinea = Nothing
    End If
    
    If LINEA_ Then
        ReDim xCampos(3, 4) As String
        
        'descripcion                            'campo                          'tamaño                        'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripcion":          xCampos(0, 1) = "descline":     xCampos(0, 2) = "4000":        xCampos(0, 3) = "C"
        xCampos(1, 0) = "Operarios":            xCampos(1, 1) = "numop":        xCampos(1, 2) = "1000":        xCampos(1, 3) = "N"
        xCampos(2, 0) = "Eficiencia (%)":       xCampos(2, 1) = "efic":         xCampos(2, 2) = "1250":        xCampos(2, 3) = "N"
     
        cSQL = "SELECT pro_linea.descripcion AS descline, pro_linea.numop, pro_linea.efic, pro_linea.idlinea, pro_linea.id AS idlineadet " _
            + vbCr + "From pro_linea " _
            + vbCr + "WHERE (((pro_linea.idrec)=" & NulosN(lblIdRec.Caption) & "));"
    
        nTitulo = "Buscando Linea"
        CARGAR_DLL_EPSBUSCAR xCon, RstTmp, cSQL, xCampos(), nTitulo, "descline", "descline", Principio
    
        If RstTmp.State = 0 Then Exit Sub
        If RstTmp.RecordCount = 0 Then Exit Sub
        ' Se filtran las tareas y Personal Involucrados
        RstTareasAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) & ""
        RstPersonalAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) & ""
        
        Dim MOSTRARMENSAJE As Boolean
        
        MOSTRARMENSAJE = False
        If RstTareasAux.RecordCount <> 0 And RstPersonalAux.RecordCount <> 0 Then MOSTRARMENSAJE = True
        If MOSTRARMENSAJE Then
            Rpta = MsgBox("¿Se Eliminara Todo el Personal y Tareas Relacionado a la linea Anterior; desea continuar?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbNo Then Exit Sub
        End If
        
        fg(0).Rows = fg(0).FixedRows
        ' Se Limpia las Tareas
        limpiarRST RstTareasAux, False
        ' Se Limpia el personal
        limpiarRST RstPersonalAux, False
        
        
        ' Se llenan los Datos de la linea
        TxtIdLineaDet.Text = NulosN(RstTmp("idlineadet"))
        LblLinea.Caption = NulosC(RstTmp("descline"))
        
        Cmd(2).SetFocus
        
        Set RstTmp = Nothing
    End If
End Sub

Private Sub pCargarDatos(fGrid As VSFlexGrid, _
                        Optional PERSONAL_ As Boolean = True, _
                        Optional TAREAS_ As Boolean = False, _
                        Optional TODOS_ As Boolean = False, _
                        Optional RECETA_ As Boolean = False, _
                        Optional NUEVO_ As Boolean = True)
    
    Dim A As Integer
    
    Agregando = True
        
    With fGrid
        If PERSONAL_ Then ' Si se desea cargar personal
            .Rows = 1
            If RstPersonal.State = 0 Then Exit Sub
            
            If NUEVO_ Then
                RstPersonal.Filter = adFilterNone
                RstPersonal.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) & ""
                ' Se verifica que este creado el recordset
                If RstPersonalAux.State = 0 Then DEFINIR_RST_TMP RstPersonalAux, RstPersonal
                ' Se vacia el recordset
                limpiarRST RstPersonalAux
                ' Se carga con los datos temporales
                CARGAR_RST_TMP RstPersonalAux, RstPersonal
            End If
            
            If TODOS_ Then ' si se muestran todos los trabajadores
                RstPersonalAux.Filter = adFilterNone
                RstPersonalAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) & ""
            Else ' si se muestran solo de una tarea especifica
                RstPersonalAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) _
                                            & " And idtar = " & NulosN(LblIdTarea.Caption) & ""
            End If
            
            If RstPersonalAux.RecordCount = 0 Then ' Si no hay Personal
                LblDetTrab.Caption = (.Rows - 1) & " de " & NulosN(LblNTrab.Caption)
                Exit Sub
            End If
            ' Se llena al Personal
            Do While Not RstPersonalAux.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = NulosN(RstPersonalAux.Fields("activo"))
                .TextMatrix(.Rows - 1, 2) = NulosC(RstPersonalAux.Fields("codigo"))
                .TextMatrix(.Rows - 1, 3) = NulosC(RstPersonalAux.Fields("nombre"))
                .TextMatrix(.Rows - 1, 4) = NulosN(RstPersonalAux.Fields("idcrdet"))
                .TextMatrix(.Rows - 1, 5) = NulosN(RstPersonalAux.Fields("idper"))
                .TextMatrix(.Rows - 1, 6) = NulosN(RstPersonalAux.Fields("idtar"))
                .TextMatrix(.Rows - 1, 7) = NulosC(RstPersonalAux.Fields("destar"))
                
                RstPersonalAux.MoveNext
            Loop
            ' Se llena el numero de trabajadores
            If TODOS_ Then
                LblDetTrab.Caption = .Rows - 1
            Else
                LblDetTrab.Caption = .Rows - 1 & " de " & NulosN(LblNTrab.Caption)
            End If
            ' aplicando el orden a la lista de datos
            GRID_ORDENAR fGrid, 1, 2
        End If
        
        If TAREAS_ Then ' Si se desea cargar Tareas
            .Rows = 1
            ' Si no hay Tareas guardadas
            If RstTareas.State = 0 Then Exit Sub
            ' Se verfica si es una carga nueva o actualizacion de datos
            If NUEVO_ Then
                ' Se filtra el registro involucrado
                RstTareas.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) & ""
                If RstTareasAux.State = 0 Then DEFINIR_RST_TMP RstTareasAux, RstTareas
                limpiarRST RstTareasAux
                CARGAR_RST_TMP RstTareasAux, RstTareas
            End If
            
            RstTareasAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) & ""
            If RstTareasAux.RecordCount = 0 Then Exit Sub
            
            RstTareasAux.MoveFirst
            ' Se llena el id y el detalle de la Linea y del encargado
            TxtIdLineaDet.Text = NulosN(RstTareasAux.Fields("idlinea"))
            TxtIdLineaDet_Validate True
            TxtIdEncarg.Text = NulosC(RstTareasAux.Fields("idresp"))
            LblEncargado.Caption = NulosC(RstTareasAux.Fields("nomresp"))
            ' Se procede a llenar las tareas
            Do While Not RstTareasAux.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = NulosN(RstTareasAux.Fields("activo"))
                .TextMatrix(.Rows - 1, 2) = NulosC(RstTareasAux.Fields("destar"))
                .TextMatrix(.Rows - 1, 3) = Format(RstTareasAux.Fields("durtar"), "HH:mm")
                .TextMatrix(.Rows - 1, 4) = Format(RstTareasAux.Fields("horinitar"), "HH:mm")
                .TextMatrix(.Rows - 1, 5) = Format(RstTareasAux.Fields("horfintar"), "HH:mm")
                .TextMatrix(.Rows - 1, 6) = Format(NulosN(RstTareasAux.Fields("numper")), "00")
                .TextMatrix(.Rows - 1, 7) = Format(NulosN(RstTareasAux.Fields("cantproc")), "0.00")
                .TextMatrix(.Rows - 1, 8) = NulosC(RstTareasAux.Fields("fchini"))
                .TextMatrix(.Rows - 1, 9) = NulosC(RstTareasAux.Fields("fchfin"))
                .TextMatrix(.Rows - 1, 10) = NulosN(RstTareasAux.Fields("idcrdet"))
                .TextMatrix(.Rows - 1, 11) = NulosN(RstTareasAux.Fields("idtar"))
                .TextMatrix(.Rows - 1, 12) = NulosN(RstTareasAux.Fields("aplpor"))
                
                If NulosN(RstTareasAux.Fields("activo")) = True Then
                    fg(0).Select .Rows - 1, 1, .Rows - 1, .Cols - 1
                    fg(0).FillStyle = flexFillRepeat
                    fg(0).CellBackColor = &HFFFF&     '&H8000000F&
                End If
                
                RstTareasAux.MoveNext
            Loop
            fg(0).Select 1, 1
            ' Si no hay registros
            If .Rows = 1 Then
                DTPHoraFin.Value = 0
                TxtCantMP.Text = ""
            Else
                ' Se verfica que la primera tarea tenga porcentaje 100% para evitar malos calculos
                If .TextMatrix(1, 12) = 0 Then
                    .TextMatrix(1, 12) = 100
                End If
                ' Se Busca la Ultima Tarea Seleccionada
                ' para llenar la fecha y hora de fin
                For A = .Rows - 1 To 1 Step -1
                    If .TextMatrix(A, 1) = 0 Then GoTo SIGUIENTEULTIMO
                    DTPHoraFin.Value = Format(.TextMatrix(A, 5), "HH:mm")
                    ' Se Filtra la Tarea
                    RstTareasAux.Filter = "idcrdet = " & NulosN(.TextMatrix(A, 10)) & _
                                                    " And idtar = " & NulosN(.TextMatrix(A, 11)) & ""
                    If RstTareasAux.RecordCount <> 0 Then TxtBFchFin.valor = RstTareasAux("fchfin")
                    A = 1
SIGUIENTEULTIMO:
                Next A
                
                ' Se Busca la primera Tarea Seleccionada
                ' para llenar la cantidad de Mp
                For A = 1 To .Rows - 1
                    If .TextMatrix(A, 1) = 0 Then GoTo SIGUIENTEPRIMERO
                    TxtCantMP.Text = Format((.TextMatrix(A, 7) * 100) / .TextMatrix(A, 12), "0.00")
                    A = .Rows - 1
SIGUIENTEPRIMERO:
                Next A
            End If
            
        End If
        
        If RECETA_ Then ' No disponible
        End If
    End With
    
    Agregando = False
End Sub

Private Sub procesarCronograma(RstTareas_Aux As ADODB.Recordset, _
                        Optional es_nuevo As Boolean = True, _
                        Optional cantidad_procesada As Double = 0, _
                        Optional hora_inicio As String = "00:00", _
                        Optional hora_fin As String = "00:00", _
                        Optional fecha_fin As Date = "25/05/2011", _
                        Optional IDITEM_ As Double = 0, _
                        Optional ID_CRDET_ As Double = 0, _
                        Optional IDLINEADET_ As Double = 0, _
                        Optional IDRESPONSABLE_ As Integer = 0, _
                        Optional NOMRESPONSABLE_ As String = "")

    Dim xTiempo As Double               ' duracion de tarea en formato numero
    Dim xHorEst As String               ' duracion de tarea en formato HH:mm
    Dim fecha_Inicio_Tarea As Date
    Dim fecha_fin_tarea As Date
    Dim CANTIDAD_ As Double
    Dim A, B As Integer
    
    Dim cantidad_procesada_anterior As Double
    Dim hora_inicio_tarea_anterior As String
    Dim hora_fin_tarea_anterior As String
    Dim duracion_tarea_anterior As String
    
    Dim Tipo As Integer
    Dim valor As Variant
    Dim considerar_refrigerio As Boolean
    Dim hor_ini_refrigerio As String
    Dim hor_fin_refrigerio As String
    
    ' Se dan los valores segun Opciones
    Tipo = MODO_TAREA
    If Tipo = 2 Then valor = NulosC(MINUTOS_) Else valor = NulosN(PORCENTAJE)
    If INCLUIR_HORAS = 0 Then considerar_refrigerio = True Else considerar_refrigerio = False
    hor_ini_refrigerio = HOR_INI
    hor_fin_refrigerio = HOR_FIN
    
    ' Si el responsable es vacio se llena con el valor por defecto
    If IDRESPONSABLE_ = 0 Then IDRESPONSABLE_ = NulosN(TxtIdEncarg.Text)
    If NOMRESPONSABLE_ = "" Then NOMRESPONSABLE_ = NulosC(LblEncargado.Caption)

    If RstTareas_Aux.State = 0 Then Exit Sub
    If RstTareas_Aux.RecordCount = 0 Then Exit Sub
    
    RstTareas_Aux.MoveFirst
    
    Agregando = True
    
    Dim xRs As New ADODB.Recordset
    
    DEFINIR_RST_TMP xRs, RstTareas_Aux
    CARGAR_RST_TMP xRs, RstTareas_Aux
    
    RstTareas_Aux.MoveFirst
    
    ' Se halla los valores iniciales de los campos cuando no es un ingreso nuevo
    If Not es_nuevo Then
        cantidad_procesada_anterior = cantidad_procesada
        hora_inicio_tarea_anterior = hora_inicio
        hora_fin_tarea_anterior = hora_fin
        duracion_tarea_anterior = Format(CDate(hora_fin) - CDate(hora_inicio), "HH:mm")
        
        fecha_fin_tarea = fecha_fin
        fecha_Inicio_Tarea = fecha_fin
    End If
    
    ' Se proceden a procesar y agregar todos los productos filtrados
    For B = 1 To RstTareas_Aux.RecordCount
        If RstTareas_Aux("activo") = 0 Then GoTo SIGUIENTE
        
        If es_nuevo Then ' Si es nuevo se agrega un nuevo registro al Recordset de Tareas
            RstTareasAux.AddNew
        Else ' Sino se filtra el registro involucrado
            RstTareasAux.Filter = "idcrdet = " & ID_CRDET_ & " And idtar = " & RstTareas_Aux("idtar") & ""
        End If
        
        ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>Se llena la cantidad de producto
        If B = 1 And es_nuevo Then
            ' Se calcula segun el rendimiento la cantidad a necesitar
            xRs.Filter = adFilterNone
            xRs.Filter = "idcrdet = " & ID_CRDET_ & ""
            
            CANTIDAD_ = calcularRdmto(NulosN(TxtIdLineaDet.Text), ID_CRDET_, xRs, NulosN(RstTareas_Aux("cantidad")))
        Else
            ' En las siguientes tareas la cantidad es la procesada en la tarea anterior
            CANTIDAD_ = NulosN(cantidad_procesada_anterior)
            
        End If

        ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>Se llena la cantidad porcentual
        ' Si el procentaje aplicado a la tarea es cero
        Dim PORCENTAJE_AUX As Double
        PORCENTAJE_AUX = NulosN(RstTareas_Aux("aplpor"))
        If PORCENTAJE_AUX = 0 Then PORCENTAJE_AUX = 100
        
        RstTareasAux("cantproc") = (NulosN(CANTIDAD_) * ((PORCENTAJE_AUX / 100)))
        
        cantidad_procesada_anterior = RstTareasAux("cantproc")

        ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>Se calcula el tiempo de demora de la tarea
        xTiempo = 0
        Dim FACTOR_ As Double
        Dim HORARR_ As Double
        Dim RstFactor As New ADODB.Recordset
        
        Set RstFactor = Nothing
        
        cSQL = "SELECT pro_lineadet.factor, pro_lineadet.intervalo AS horarr " _
            + vbCr + "From pro_lineadet " _
            + vbCr + "Where (((pro_lineadet.idlineadet) = " & IDLINEADET_ & ") And ((pro_lineadet.IDTAR) = " & NulosN(RstTareas_Aux("idtar")) & ")) " _
            + vbCr + "GROUP BY pro_lineadet.factor, pro_lineadet.intervalo;"
            
        RST_Busq RstFactor, cSQL, xCon
        
        If RstFactor.State = 0 Then FACTOR_ = 0: HORARR_ = 0
        If RstFactor.RecordCount = 0 Then FACTOR_ = 0 Else FACTOR_ = NulosN(RstFactor("factor")): HORARR_ = NulosN(RstFactor("horarr"))
        
        If NulosN(RstTareas_Aux("numper")) <> 0 Then
            xTiempo = (FACTOR_ * CANTIDAD_) / NulosN(RstTareas_Aux("numper"))
        End If
        If xTiempo > 24 Then xTiempo = 23.9
        ' Tiempo de duracion de la tarea
        xHorEst = ""
        xHorEst = Format(Int(xTiempo), "00")
        xHorEst = xHorEst & ":" & Format(((xTiempo * 60) Mod 60), "00")
        RstTareasAux("durtar") = xHorEst

        ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>Se llena la hora de inicio de la tarea
        Dim h() As String
        Dim tiempo As Double
        Dim intervalo As String

        If B = 1 And es_nuevo Then ' Si es el primer ingreso nuevo
            RstTareasAux("horinitar") = RstTareas_Aux("horinitar")
        Else
            If Tipo = 0 Then ' una tarea despues de otra
                RstTareasAux("horinitar") = hora_fin_tarea_anterior
            End If

            If Tipo = 1 Then ' una tarea al porcentaje de otra
                ' Se aplica el porcentaje
                h = Split(duracion_tarea_anterior, ":")
                tiempo = (60 * Val(h(0))) + Val(h(1))
                tiempo = ((valor * tiempo) / 100)
                tiempo = tiempo / 60 ' Se cambia a horas

                intervalo = Format(Int(tiempo), "00")
                intervalo = intervalo & ":" & Format(((tiempo * 60) Mod 60), "00")
                RstTareasAux("horinitar") = CDate(hora_inicio_tarea_anterior) + CDate(intervalo)
            End If

            If Tipo = 2 Then ' Una tarea al minuto de otra
                If hora_inicio_tarea_anterior = hora_fin_tarea_anterior Then
                    RstTareasAux("horinitar") = hora_inicio_tarea_anterior
                Else
                    RstTareasAux("horinitar") = CDate(hora_inicio_tarea_anterior) + CDate(valor)
                End If
            End If
            
            If Tipo = 3 Then ' Segun Receta
                If hora_inicio_tarea_anterior = hora_fin_tarea_anterior Then
                    RstTareasAux("horinitar") = hora_inicio_tarea_anterior
                Else
                    intervalo = Format(Int(HORARR_), "00")
                    intervalo = intervalo & ":" & Format(((HORARR_ * 60) Mod 60), "00")
                    RstTareasAux("horinitar") = CDate(hora_inicio_tarea_anterior) + CDate(intervalo)
                End If
            End If
        End If

        If considerar_refrigerio Then ' Considerar horarios de refrigerio
            ' Si la hora de inicio de la tarea esta entre los horarios de refrigerio
            ' La hora de inicio es el del fin de refrigerio
            If (RstTareasAux("horinitar") > CDate(hor_ini_refrigerio)) And (RstTareasAux("horinitar") < CDate(hor_fin_refrigerio)) Then
                RstTareasAux("horinitar") = CDate(hor_fin_refrigerio)
            End If
        End If
        
        duracion_tarea_anterior = Format(RstTareasAux("durtar"), "HH:mm")
        hora_inicio_tarea_anterior = Format(RstTareasAux("horinitar"), "HH:mm")

        ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>Se llena la fecha de inicio de la tarea
        If B = 1 And es_nuevo Then
            fecha_Inicio_Tarea = CDate(Format(RstTareas_Aux("fchini"), "dd/mm/yy") & " " & Format(RstTareas_Aux("horinitar"), "HH:mm"))
        Else
            'fecha_Inicio_Tarea = CDate(Format(fecha_fin_tarea, "dd/mm/yy") & " " & Format(RstTareasAux("horinitar"), "HH:mm")) '+ CDate(Fg2.TextMatrix(A, 16))
            fecha_Inicio_Tarea = CDate(Format(fecha_Inicio_Tarea, "dd/mm/yy") & " " & Format(RstTareasAux("horinitar"), "HH:mm")) '+ CDate(Fg2.TextMatrix(A, 16))
        End If
        RstTareasAux("fchini") = Format(fecha_Inicio_Tarea, "dd/mm/yy") ' fecha de inicio de la tarea

        ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>Se llena la fecha de fin de la tarea
        fecha_fin_tarea = fecha_Inicio_Tarea + RstTareasAux("durtar")
        RstTareasAux("fchfin") = Format(fecha_fin_tarea, "dd/mm/yy") ' fecha de fin de la tarea
        
        ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>Se llena la hora de fin de la tarea
        RstTareasAux("horfintar") = Format(fecha_fin_tarea, "HH:mm")

        If considerar_refrigerio Then ' Considerar horarios de refrigerio
            Dim durac_ref As String
            durac_ref = Format(CDate(hor_fin_refrigerio) - CDate(hor_ini_refrigerio), "HH:mm")

            ' Si la hora de fin de la tarea esta entre los horarios de refrigerio
            ' Se aumenta a la hora de fin la duracion de la tarea
            If (RstTareasAux("horfintar") > CDate(hor_ini_refrigerio)) And (RstTareasAux("horfintar") <= CDate(hor_fin_refrigerio)) Then
                RstTareasAux("horfintar") = RstTareasAux("horfintar") + CDate(durac_ref)
            Else
                ' Si el refrigerio esta entre la hora de inicio y fin de la tarea
                ' Se aumenta a la hora de fin la duracion de la tarea
                If (RstTareasAux("horinitar") <= CDate(hor_ini_refrigerio)) And (RstTareasAux("horfintar") >= CDate(hor_fin_refrigerio)) Then
                    RstTareasAux("horfintar") = RstTareasAux("horfintar") + CDate(durac_ref)
                End If
            End If
        End If
         
        If ID_CRDET_ = 0 Then
            RstTareasAux("idcrdet") = NulosN(LblIdCrDet.Caption)
        Else
            RstTareasAux("idcrdet") = NulosN(ID_CRDET_)
        End If
        
        If es_nuevo Then RstTareasAux("activo") = -1
        
        RstTareasAux("idtar") = NulosN(RstTareas_Aux("idtar"))                        ' id de tarea
        RstTareasAux("orden") = NulosN(RstTareas_Aux("orden"))                        ' Orden de la tarea
        RstTareasAux("destar") = NulosC(RstTareas_Aux("destar"))                      ' nombre de la tarea
        RstTareasAux("numper") = Format(NulosN(RstTareas_Aux("numper")), "00")        ' numero de personas para la tarea
        RstTareasAux("aplpor") = Format(NulosN(RstTareas_Aux("aplpor")), "0.00")      ' rendimiento para la cantidad de producto
        RstTareasAux("idlinea") = IDLINEADET_                                         ' rendimiento para la cantidad de producto
        RstTareasAux("idresp") = IDRESPONSABLE_                                       ' codigo del responsable
        RstTareasAux("nomresp") = NOMRESPONSABLE_                                     ' nombre del responsable
        
        RstTareasAux.Update
        
SIGUIENTE:
        RstTareas_Aux.MoveNext
        
        If RstTareas_Aux.EOF = True Then
            Exit For
        End If
    Next B
        
    Agregando = False
End Sub

Private Function calcularHoraFin(IDITEM As Integer, FECHA_DE_INICIO As Date, CANTIDAD As Double) As Date
    Dim RstLinea As New ADODB.Recordset
    Dim xTiempo As Double
    Dim xHorEst As String
    
    cSQL = "SELECT pro_receta.id, pro_receta.iditem, pro_recetalinea.idunimed, pro_recetalinea.frechora " _
            + vbCr + "FROM pro_receta RIGHT JOIN pro_recetalinea ON pro_receta.id = pro_recetalinea.idrec " _
            + vbCr + "Where (((pro_receta.prirec) = 1)) " _
            + vbCr + "GROUP BY pro_receta.id, pro_receta.iditem, pro_recetalinea.idunimed, pro_recetalinea.frechora " _
            + vbCr + "HAVING (((pro_receta.iditem)=" & IDITEM & "));"
            
    RST_Busq RstLinea, cSQL, xCon
    
    If RstLinea.State = 0 Then GoTo ERROR_AL_ENCONTRAR_LINEA
    If RstLinea.RecordCount = 0 Then GoTo ERROR_AL_ENCONTRAR_LINEA
    

    xTiempo = NulosN(CANTIDAD) / NulosN(RstLinea("frechora"))
    xHorEst = Format(Int(xTiempo), "00")
    xHorEst = xHorEst & ":" & Format(((xTiempo * 60) Mod 60), "00")
    
    calcularHoraFin = FECHA_DE_INICIO + CDate(xHorEst)
    Exit Function
    
ERROR_AL_ENCONTRAR_LINEA:
    MsgBox "No se ha podido procesar el tiempo final para este Producto, verifique si tiene una linea activa que procesar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    calcularHoraFin = FECHA_DE_INICIO
End Function

Function DateFromString(DatePart As String, TimePart As String) As Date
    Dim dtDatePart As Date, dtTimePart As Date
    dtDatePart = DatePart
    dtTimePart = TimePart
    DateFromString = dtDatePart + dtTimePart
End Function

Private Sub CmdCancelar_Click(Index As Integer)
    If Index = 0 Then
        TabOne1.Enabled = True
        Toolbar1.Enabled = True
        
        CmdAcepta(0).Enabled = True
        visEvent = False
        Frame3.Visible = False
    End If
End Sub

Private Sub CmdOpciones_Click(Index As Integer)
    Dim xFrm As New sgi2_produccion.produccion
    
    If Index = 5 Then ' Cambiar Vista
        If CalendarControl1.ViewType = xtpCalendarDayView Then
            CalendarControl1.ViewType = xtpCalendarFullWeekView
        Else
            If CalendarControl1.ViewType = xtpCalendarFullWeekView Then
                CalendarControl1.ViewType = xtpCalendarDayView
            End If
        End If
    End If
    
    If QueHace = 3 Then Exit Sub
    
    If Index = 0 Then ' Procesar
        If TxtFchIni.valor = "" Then
            MsgBox "No ha especificado la fecha de inicio", vbInformation + vbOKOnly + vbDefaultButton1
            TxtFchIni.SetFocus
            Exit Sub
        End If
    
        If TxtFchFin.valor = "" Then
            MsgBox "No ha especificado la fecha final", vbInformation + vbOKOnly + vbDefaultButton1
            TxtFchFin.SetFocus
            Exit Sub
        End If
    
        If NulosN(TxtTipPro.Text) = 0 Then
            MsgBox "No ha especificado el tipo de producto a procesar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtTipPro.SetFocus
            Exit Sub
        End If
        
        If QueHace = 1 Then CalendarControl1.Visible = True: CmdOpciones(0).Enabled = False
        CalendarControl1.ActiveView.ShowDay (CDate(TxtFchIni.valor))
        CalendarControl1.ViewType = xtpCalendarFullWeekView
        
        TxtIdSup.Locked = True
        ComboSemanas.Locked = True
        TxtFchIni.Locked = True
        TxtFchFin.Locked = True
        CmdBusTip.Enabled = False
        TxtTipPro.Locked = True
        
        CmdOpciones(1).Enabled = True
        CmdOpciones(2).Enabled = True
        CmdOpciones(3).Enabled = True
        '****************************************************************************
        CmdOpciones(4).Enabled = True
        '****************************************************************************
               
        If CalendarControl1.Visible = True Then CalendarControl1.SetFocus
    End If
    
    If Index = 1 Then ' Agregar
        mostrarFormulario
    End If
    
    If Index = 2 Then ' Modificar
        mostrarFormulario False, True, False
    End If
    
    If Index = 3 Then ' Eliminar
        menu2_2_Click
    End If
    
    If Index = 4 Then ' Linea de Produccion
        ' Se llama al formulario de linea de produccion
        xFrm.CronogramaPlaneaProduccion xCon
        Set xFrm = Nothing
    End If
End Sub

Private Sub ComboSemanas_Click()
    If QueHace <> 3 Then
        Dim fechaI As Date
        Dim fechaF As Date
        calcularSemana ComboSemanas.Text, fechaI, fechaF
        CAMBIO_ = True
        TxtFchIni.valor = fechaI
        TxtFchFin.valor = fechaF
        CAMBIO_ = False
        If CmdBusTip.Enabled = True Then CmdBusTip.SetFocus
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : LlenarDatos
'* Tipo             : SUB
'* Descripcion      : CARGA LOS DATOS AL CALENDARIO
'* Modificacion     : 15/02/11 JOSE CHACON
'*                      21/04/2011 -> se modifica la referencia "id" de pro_cronogramadet por "idcr"
'*****************************************************************************************************
Sub LlenarDatos()
    Dim EVENTONUEVO_ As CalendarEvent
    Dim A As Integer
    Dim xRs As New ADODB.Recordset
    
    ' Se llenan los Productos
    If RstProductos.State = 0 Then
        llenarDefinirRST NulosN(RstLis("semana")), False, False, False, True
    End If
    
    If RstProductos.State = 0 Then Exit Sub
    If RstProductos.RecordCount = 0 Then Exit Sub
    
    DEFINIR_RST_TMP xRs, RstProductos
    CARGAR_RST_TMP xRs, RstProductos
    
    'se crea un evento nuevo de calendario
    Set EVENTONUEVO_ = CalendarControl1.DataProvider.CreateEvent
    
    'se procede a llenar los detalles del evento
    xRs.MoveFirst
    For A = 1 To xRs.RecordCount
        EVENTONUEVO_.ScheduleID = NulosN(xRs("iditem"))
        EVENTONUEVO_.Subject = NulosC(xRs("descripcion"))
        EVENTONUEVO_.StartTime = Format(xRs("fchpro"), "dd/mm/yyyy") & " " & NulosC(Format(xRs("horpro"), "HH:mm"))
        EVENTONUEVO_.EndTime = Format(xRs("fchfin"), "dd/mm/yyyy") & " " & NulosC(Format(xRs("horfin"), "HH:mm"))
        EVENTONUEVO_.Location = NulosC(xRs("numprod"))
        EVENTONUEVO_.Body = NulosC(xRs("cantidad")) & " " & NulosC(xRs("abrev")) & _
                                    vbCr + NulosC(Format(xRs("horpro"), "HH:mm")) & " - " _
                                    & NulosC(Format(xRs("horfin"), "HH:mm"))
        EVENTONUEVO_.ReminderSoundFile = NulosC(xRs("id"))
        'se agrega el evento nuevo al calendario
        CalendarControl1.DataProvider.AddEvent EVENTONUEVO_
        
        xRs.MoveNext
    Next A
    
    Set xRs = Nothing
End Sub

Sub calcularSemana(numSemana As Integer, ByRef fechaInicio As Date, ByRef fechaFin As Date)
    Dim fechaRef As Date
    fechaRef = CDate("01/01/" & AnoTra)
    
    'Buscamos el primer Lunes del Año
    While Weekday(fechaRef) <> vbMonday
        'Vamos sumando dia a dia, hasta encontrar el primer lunes
        fechaRef = fechaRef + 1
    Wend
    
    'Multiplicamos y obtenemos el rango inferior de la semana
    fechaInicio = fechaRef + (7 * (numSemana - 1))
    'Obtenemos el rango superior de la semana
    fechaFin = fechaInicio + 6
End Sub

'*****************************************************************************************************
'* Nombre           : CmAcepta_Click
'* Tipo             : SUB
'* Descripcion      :
'* Modificacion     : 15/02/11 JOSE CHACON
'*                      21/04/2011 -> se modifica la referencia "id" de RstMatPro por "idcr"
'*****************************************************************************************************
Private Sub CmdBusSup_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Apellidos y Nombres":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":                xCampos(1, 1) = "id":        xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    cSQL = "SELECT pro_emp.*, pla_empleados.nombre " _
        + vbCr + "FROM (pro_emp LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id " _
        + vbCr + "Where (((pro_empdet.idfun) = 2)) " _
        + vbCr + "ORDER BY pla_empleados.nombre;"
    
    xform.SQLCad = cSQL
    
    xform.titulo = "Buscando Supervisores"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdSup.Text = xRs("id")
            LblSupervisor.Caption = xRs("nombre")
            If CmdOpciones(0).Enabled = True Then
                CmdOpciones(0).SetFocus
            Else
                CalendarControl1.SetFocus
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

Private Sub CmdBusTip_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion                     'campo                            'tamaño                         'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_tipoproducto.id, mae_tipoproducto.descripcion FROM mae_tipoproducto"
    
    xform.titulo = "Buscando Tipo de Item"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipPro.Text = xRs("id")
            LblTipoProd.Caption = xRs("descripcion")
            CmdOpciones(0).SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : Function
'* Descripcion      : GRABA LOS DATOS DEL CALENDARIO
'* Modificacion     : 15/02/11 JOSE CHACON
'*                      21/04/2011 -> se modifica la referencia "id" de pro_cronogramadet por "idcr"
'*                      21/04/2011 -> se agrega "identificador" para grabar el id de cronogramadet y cronogramadetprod
'*****************************************************************************************************
Function Grabar() As Boolean
    Dim A, B As Integer
    Dim xTot As Long
    Dim IDCRDET_ As Double
    Dim IDCRTAR_ As Double
    Dim IDCRPERS_ As Double
    Dim IDORD_ As Double
    Dim IDORDDET_ As Double
    Dim NUMSOLIC_ As Double
    
    Dim RstSolMat As New ADODB.Recordset
    Dim xIdSol As Double
    Dim RstSolMatDet As New ADODB.Recordset
    Dim numDoc As Double
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstDet1 As New ADODB.Recordset
    
    Dim RstPers As New ADODB.Recordset
    Dim RstTar As New ADODB.Recordset
    Dim RstOrd As New ADODB.Recordset
    Dim RstOrdDet As New ADODB.Recordset
    Dim RstOrdIns As New ADODB.Recordset
    
    Dim xRs As New ADODB.Recordset
    
    Dim xId As Double
    Dim nSQL As String
    
    Dim pEvent As CalendarEvent
    Dim Events As CalendarEvents
    
    ' VERIFICAMOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
    If TxtIdSup.Text = "" Then
        MsgBox "No ha especificado un Supervisor para el nuevo Cronograma, especifique uno", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        CmdBusSup.SetFocus
        Exit Function
    End If
    
    If ComboSemanas.Text = "" Then
        MsgBox "No ha especificado una fecha para el Cronograma, especifique una", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        ComboSemanas.SetFocus
        Exit Function
    End If
    
    If TxtTipPro.Text = "" Then
        MsgBox "No ha especificado un tipo de Producto para el Cronograma, especifique uno", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        CmdBusTip.SetFocus
        Exit Function
    End If
    
On Error GoTo LaCague
    xCon.BeginTrans

    If QueHace = 1 Then
        ' SI ES UN NUEVO REGISTRO OBTENEMOS EL ULTIMO ID DE LA TABLA
        xId = HallaCodigoTabla("pro_cronograma", xCon, "id")
        
        mIdRegistro = NulosN(ComboSemanas.Text)
    Else
        'Busco todos los cronogramas relacionados con esa semana
        cSQL = "SELECT pro_cronograma.id AS idcr, pro_cronograma.semana " _
            + vbCr + "From pro_cronograma " _
            + vbCr + "Where (((pro_cronograma.semana) = " & NulosN(ComboSemanas.Text) & ")) " _
            + vbCr + "GROUP BY pro_cronograma.id, pro_cronograma.semana;"
        
        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 0 Then GoTo LaCague
        If xRs.RecordCount = 0 Then GoTo LaCague
        
        xRs.MoveFirst
        While Not xRs.EOF
            xId = NulosN(xRs("idcr"))
            ' Eliminamos los registros involucrados
            ' Ordenes de Produccion
            xCon.Execute "DELETE * FROM pro_ordenproddetins WHERE idcr  = " & xId & ""
            xCon.Execute "DELETE * FROM pro_ordenproddet WHERE idcr  = " & xId & ""
            xCon.Execute "DELETE * FROM pro_ordenprod WHERE idcr  = " & xId & ""
            ' Personal de Produccion
            xCon.Execute "DELETE * FROM pro_cronogramapers WHERE idcr  = " & xId & ""
            ' Tareas de Produccion
            xCon.Execute "DELETE * FROM pro_cronogramatarea WHERE idcr  = " & xId & ""
            ' Detalle
            xCon.Execute "DELETE * FROM pro_cronogramadetprod WHERE idcr  = " & xId & ""
            xCon.Execute "DELETE * FROM pro_cronogramadet WHERE idcr  = " & xId & ""
            ' Cabecera
            xCon.Execute "DELETE * FROM pro_cronograma WHERE id  = " & xId & ""
            
            xRs.MoveNext
        Wend
        
        mIdRegistro = RstLis("semana")
    End If
    ' Cabecera
    RST_Busq RstCab, "SELECT TOP 1 * FROM pro_cronograma", xCon
    ' Detalle
    RST_Busq RstDet, "SELECT TOP 1 * FROM pro_cronogramadet", xCon
    RST_Busq RstDet1, "SELECT TOP 1 * FROM pro_cronogramadetprod", xCon
    ' Personas
    RST_Busq RstPers, "SELECT TOP 1 * FROM pro_cronogramapers", xCon
    'Tareas
    RST_Busq RstTar, "SELECT TOP 1 * FROM pro_cronogramatarea", xCon
    'Orden
    RST_Busq RstOrdIns, "SELECT TOP 1 * FROM pro_ordenproddetins", xCon
    RST_Busq RstOrdDet, "SELECT TOP 1 * FROM pro_ordenproddet", xCon
    RST_Busq RstOrd, "SELECT TOP 1 * FROM pro_ordenprod", xCon
        
    ' SE LLENA LA CABECERA
    RstCab.AddNew
    RstCab("id") = xId
    RstCab("idsup") = NulosC(TxtIdSup.Text)
    RstCab("fchini") = NulosC(TxtFchIni.valor)
    RstCab("fchfin") = NulosC(TxtFchFin.valor)
    RstCab("idtippro") = NulosN(TxtTipPro.Text)
    RstCab("semana") = NulosN(ComboSemanas.Text)
    RstCab.Update
    
    Set Events = CalendarControl1.DataProvider.GetAllEventsRaw
        
    IDCRDET_ = HallaCodigoTabla("pro_cronogramadet", xCon, "id")
    IDCRPERS_ = HallaCodigoTabla("pro_cronogramapers", xCon, "id")
    IDCRTAR_ = HallaCodigoTabla("pro_cronogramatarea", xCon, "id")
    IDORD_ = HallaCodigoTabla("pro_ordenprod", xCon, "id")
    IDORDDET_ = HallaCodigoTabla("pro_ordenproddet", xCon, "id")
    NUMSOLIC_ = HallaCodigoTabla("pro_ordenproddet", xCon, "numdoc")
    
    RstProductos.Filter = adFilterNone
    RstProductos.MoveFirst
    For A = 1 To RstProductos.RecordCount
        Dim IDCRDETAUX_ As Double
        IDCRDETAUX_ = NulosN(RstProductos("id"))
        
        RstDet.AddNew
        RstDet("id") = IDCRDET_
        RstDet("idcr") = xId
        
        RstDet("fchpro") = NulosC(RstProductos("fchpro"))
        RstDet("fchfin") = NulosC(RstProductos("fchfin"))
        RstDet("horpro") = NulosC(Format(RstProductos("horpro"), "HH:mm"))
        RstDet("horfin") = NulosC(Format(RstProductos("horfin"), "HH:mm"))
        RstDet("iditem") = NulosN(RstProductos("iditem"))
        RstDet("idrec") = NulosN(RstProductos("idrec"))
        RstDet("cantidad") = NulosN(RstProductos("cantidad"))
        RstDet("numprod") = NulosC(RstProductos("numprod"))
        
        RstDet.Update
        
        RstPersonal.Filter = "idcrdet = " & IDCRDETAUX_ & ""
        If RstPersonal.RecordCount <> 0 Then
            RstPersonal.MoveFirst
            For B = 1 To RstPersonal.RecordCount
                RstPers.AddNew
                RstPers("id") = IDCRPERS_
                RstPers("idcr") = xId
                RstPers("idcrdet") = IDCRDET_
                
                RstPers("idper") = NulosN(RstPersonal("idper"))
                RstPers("idtar") = NulosN(RstPersonal("idtar"))
                RstPers("activo") = NulosN(RstPersonal("activo"))
                
                RstPers.Update
                IDCRPERS_ = IDCRPERS_ + 1
                RstPersonal.MoveNext
            Next B
        End If
        
        RstTareas.Filter = "idcrdet = " & IDCRDETAUX_ & ""
        If RstTareas.RecordCount <> 0 Then
            RstTareas.MoveFirst
            For B = 1 To RstTareas.RecordCount
                RstTar.AddNew
                RstTar("id") = IDCRTAR_
                RstTar("idcr") = xId
                RstTar("idcrdet") = IDCRDET_
                RstTar("idlinea") = NulosN(RstTareas("idlinea"))
                
                RstTar("idpro") = NulosN(RstProductos("iditem"))
                RstTar("fchpro") = NulosC(RstProductos("fchpro"))
                RstTar("idtar") = NulosN(RstTareas("idtar"))
                RstTar("idresp") = NulosN(RstTareas("idresp"))
                RstTar("orden") = NulosN(RstTareas("orden"))
                RstTar("activo") = NulosN(RstTareas("activo"))
                RstTar("cantproc") = NulosN(RstTareas("cantproc"))
                RstTar("numper") = NulosN(RstTareas("numper"))
                RstTar("horinitar") = NulosC(Format(RstTareas("horinitar"), "HH:mm"))
                RstTar("horfintar") = NulosC(Format(RstTareas("horfintar"), "HH:mm"))
                RstTar("durtar") = NulosC(Format(RstTareas("durtar"), "HH:mm"))
                RstTar("fchini") = NulosC(RstTareas("fchini"))
                RstTar("fchfin") = NulosC(RstTareas("fchfin"))
                RstTar("aplpor") = NulosN(RstTareas("aplpor"))
                                
                RstTar.Update
                IDCRTAR_ = IDCRTAR_ + 1
                RstTareas.MoveNext
            Next B
        End If
        IDCRDET_ = IDCRDET_ + 1
        
        RstProductos.MoveNext
    Next A
    
    ' Grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
   
    xCon.CommitTrans
    'xTitulo = "Grabar"
    MsgBox "La operacion se registró con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDet1 = Nothing
    Grabar = True
    Exit Function
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDet1 = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
End Function

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstLis
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA EN FORMA ASCENDENTE O DECENDETE LAS COLUMNAS DEL CONTROL Dg3
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstLis.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstLis("id")), xCon
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA Y DESACTIVA LOS BOTONES DE LA BARRA DE HERRAMIENTAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub DTPHoraFin_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub DTPHoras_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub Fg_Click(Index As Integer)
    Dim Rpta As Integer
    Dim ESPECIAL_ As Boolean
    
    If QueHace = 3 Then Exit Sub
    
    ESPECIAL_ = False
    If Index = 0 Then
        If fg(Index).Row < fg(Index).FixedRows Then Exit Sub
        If fg(Index).Col <> 1 Then Exit Sub
        
        If fg(Index).TextMatrix(fg(Index).Row, 1) = 0 Then ' Si se deselecciono
            'xTitulo = "Cambio en el estado de Tarea"
            Rpta = MsgBox("¿Se eliminara el Personal relacionado a esta Tarea; desea continuar?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                ' Se filtra el Personal de la Tarea y se elimina
                If fg(Index).Row > 1 Then ESPECIAL_ = True
                RstPersonalAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) & " And idtar = " & NulosN(fg(Index).TextMatrix(fg(Index).Row, 11)) & ""
                limpiarRST RstPersonalAux, False
                ' Se filtra la Tarea y se actualiza su estado
                RstTareasAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) & " And idtar = " & NulosN(fg(Index).TextMatrix(fg(Index).Row, 11)) & ""
                If RstTareasAux.RecordCount > 0 Then RstTareasAux("activo") = False
                ' Se modifica las tareas
                limpiarTarea NulosN(LblIdCrDet.Caption), NulosN(fg(Index).TextMatrix(fg(Index).Row, 11))
            Else
                ' Se selecciona de nuevo la tarea
                fg(Index).TextMatrix(fg(Index).Row, 1) = -1
            End If
        Else
            ' Se filtra el Personal de la Tarea y se elimina
            RstTareasAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) & " And idtar = " & NulosN(fg(Index).TextMatrix(fg(Index).Row, 11)) & ""
            If RstTareasAux.RecordCount > 0 Then RstTareasAux("Activo") = True
        End If
    End If
    
    If Index = 2 Then
        Dim A As Integer
        Dim contador As Integer
        
        If Frame10.Visible = False Then Exit Sub
        
        contador = 0
        For A = 1 To fg(2).Rows - 1
            If fg(2).TextMatrix(A, 1) = -1 Then contador = contador + 1
        Next A
        
        LbNumSel.Caption = Format(contador, "000")
    End If
End Sub

Private Function calcularRdmto(IDLINEADET_ As Double, IDCRDET_ As Double, RECORDSET_ As ADODB.Recordset, CANTIDADACTUAL_ As Double) As Double
    Dim xRs As New ADODB.Recordset
    Dim CANTIDAD_ As Double
    Dim RENDIMIENTO_ As Double
    Dim A As Integer
    
    cSQL = "SELECT pro_lineadet.idtar, pro_lineadet.rdmto " _
        + vbCr + "From pro_lineadet " _
        + vbCr + "Where (((pro_lineadet.idlineadet) = " & IDLINEADET_ & ")) " _
        + vbCr + "GROUP BY pro_lineadet.idtar, pro_lineadet.rdmto;"
    
    ' Se obtienen los rendimientos de todas las tareas de la linea
    RST_Busq xRs, cSQL, xCon
    If xRs.State = 0 Then RENDIMIENTO_ = 1
    If xRs.RecordCount = 0 Then RENDIMIENTO_ = 1
    With RECORDSET_
        .Filter = "idcrdet = " & IDCRDET_ & " And activo = -1"
        .MoveFirst
        RENDIMIENTO_ = 1
        For A = 1 To .RecordCount
            xRs.Filter = "idtar = " & NulosN(.Fields("idtar"))
            RENDIMIENTO_ = RENDIMIENTO_ * (NulosN(xRs("rdmto")) / 100)
            .MoveNext
        Next A
    End With
    CANTIDAD_ = CANTIDADACTUAL_ / RENDIMIENTO_
    
    calcularRdmto = CANTIDAD_
End Function

Private Function calcularProdAnterior(IDLINEADET_ As Double, IDITEM_ As Boolean, DESPROD_ As Boolean) As Variant
    Dim xRs As New ADODB.Recordset
    Dim DESCRIPCION_ As String
    Dim RENDIMIENTO_ As Double
    Dim A As Integer
    
    cSQL = "SELECT pro_recetains.iditem, alm_inventario.descripcion " _
        + vbCr + "FROM (pro_lineadet LEFT JOIN pro_recetains ON pro_lineadet.idrec = pro_recetains.idrec) LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id " _
        + vbCr + "Where (((pro_lineadet.idlineadet) = " & IDLINEADET_ & ") And ((alm_inventario.tippro) <= 3)) " _
        + vbCr + "GROUP BY pro_recetains.iditem, alm_inventario.descripcion;"
    
    ' Se obtienen los rendimientos de todas las tareas de la linea
    RST_Busq xRs, cSQL, xCon
    If xRs.State = 0 Then DESCRIPCION_ = "": GoTo SALIR
    If xRs.RecordCount = 0 Then DESCRIPCION_ = "": GoTo SALIR
    
    DESCRIPCION_ = NulosC(xRs("descripcion"))
    
SALIR:
    calcularProdAnterior = DESCRIPCION_
End Function

Private Sub limpiarTarea(IDCRDET_ As Double, IDTAR_ As Double)
    ' Se modifica la tarea seleccionada
    RstTareasAux.Filter = "idcrdet = " & IDCRDET_ & " And idtar = " & IDTAR_ & ""
    RstTareasAux("activo") = False
    RstTareasAux("durtar") = "00:00"
    RstTareasAux("horinitar") = "00:00"
    RstTareasAux("horfintar") = "00:00"
    RstTareasAux("cantproc") = 0
End Sub

Private Sub Fg_EnterCell(Index As Integer)
    If QueHace = 3 Then
        fg(Index).Editable = flexEDNone
        'fg(Index).SelectionMode = flexSelectionByRow
        Exit Sub
    End If
    fg(Index).Editable = flexEDKbdMouse
    'fg(Index).SelectionMode = flexSelectionFree
End Sub

Private Sub Fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Fg2.Col = 2 Then
        Fg2.TextMatrix(Row, Col) = Format(Fg2.TextMatrix(Row, Col), "0.00")
        
        TxtTotal.Text = Format(GRID_SUMAR_COL(Fg2, 2, 1, Fg2.Rows - 1), "0.00")
        If NulosN(Fg2.TextMatrix(Row, Col)) <> 0 Then
            Fg2.TextMatrix(Row, 3) = 1
        Else
            Fg2.TextMatrix(Row, 3) = 0
        End If
    End If
    If Fg2.Col = 3 Then
        If NulosN(Fg2.TextMatrix(Row, Col)) = 0 Then
            Fg2.TextMatrix(Row, 2) = ""
        End If
        TxtTotal.Text = Format(GRID_SUMAR_COL(Fg2, 2, 1, Fg2.Rows - 1), "0.00")
    End If
End Sub

Private Sub Fg2_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If QueHace = 3 Then KeyAscii = 0
    
    If KeyAscii = 13 Then Exit Sub
    ' validar los caracteres que se ingresan
    Select Case Col
        Case 1, 3
            KeyAscii = 0
            
        Case 2
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End Select
End Sub

Private Sub Form_Activate()
    '
    Dim Rpta As Integer
    
    If SeEjecuto = False Then
    
        SeEjecuto = True
    
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
    
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon
        
        Set RstLis = Nothing
                
        cSQL = "SELECT pro_cronograma.semana, pro_cronograma.idsup, pro_cronograma.fchini, pro_cronograma.fchfin, pro_cronograma.idtippro, mae_tipoproducto.descripcion AS destippro, pla_empleados.nombre AS apenom " _
            + vbCr + "FROM (pla_empleados RIGHT JOIN (pro_cronograma LEFT JOIN pro_emp ON pro_cronograma.idsup = pro_emp.id) ON pla_empleados.id = pro_emp.idemp) LEFT JOIN mae_tipoproducto ON pro_cronograma.idtippro = mae_tipoproducto.id " _
            + vbCr + "WHERE (((pro_cronograma.fchini) >= CDate('01/01/" & AnoTra & "'))) " _
            + vbCr + "GROUP BY pro_cronograma.semana, pro_cronograma.idsup, pro_cronograma.fchini, pro_cronograma.fchfin, pro_cronograma.idtippro, mae_tipoproducto.descripcion, pla_empleados.nombre " _
            + vbCr + "ORDER BY pro_cronograma.semana DESC;"
            
        RST_Busq RstLis, cSQL, xCon
        
        Set Dg1.DataSource = RstLis
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : Sub
'* Descripcion      : MUESTRA EL DETALLE DEL CRONOGRAMA
'* Modificacion     :
'*                    21/04/2011 JOSE CHACON
'*                      -> se modifica la referencia "id" de pro_cronogramadet por "idcr"
'*****************************************************************************************************
Sub MuestraSegundoTab()
    Dim Rpta As Integer
    Dim A As Integer
    
    If RstLis.RecordCount = 0 Then
        Rpta = MsgBox("¿No se ha encontrado ningun Cronograma creado; desea crear uno nuevo?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
        If Rpta = vbYes Then
            Nuevo
        End If
        Exit Sub
    End If
    
    TxtIdSup.Text = RstLis("idsup")
    TxtIdSup_Validate True
    TxtTipPro.Text = RstLis("idtippro")
    TxtTipPro_Validate True
    
    TxtFchIni.valor = RstLis("fchini")
    TxtFchFin.valor = RstLis("fchfin")
    
    CalendarControl1.ActiveView.ShowDay (CDate(TxtFchIni.valor))
    CalendarControl1.ViewType = xtpCalendarFullWeekView
    
    centrarFrm FraProgreso
    FraProgreso.Visible = True
    FraProgreso.Refresh
    LblProg.Caption = "PROCESANDO PRODUCTOS"
    LlenarDatos
    
    FraProgreso.Refresh
    LblProg.Caption = "PROCESANDO TAREAS"
    llenarDefinirRST NulosN(RstLis("semana")), True, False ' Tareas
    
    FraProgreso.Refresh
    LblProg.Caption = "PROCESANDO PERSONAL"
    llenarDefinirRST NulosN(RstLis("semana")), False, True ' Personal
    
'    FraProgreso.Refresh
'    LblProg.Caption = "PROCESANDO RECETA"
'    llenarDefinirRst NulosN(RstLis("semana")), False, False, True ' Receta
   
    'llenarDefinirRst NulosN(RstLis("semana")), False, False, False, False, True ' Materia Prima
    
    FraProgreso.Visible = False
    
    
    CARGO_ = True ' Se define como cargo
End Sub

Private Sub llenarDefinirRST(SEMANA_ As Double, TAREAS_ As Boolean, PERSONAL_ As Boolean, _
                                Optional RECETA_ As Boolean = False, Optional PRODUCTOS_ As Boolean = False, _
                                Optional MATERIAPRIMA_ As Boolean = False)
    Dim xRs As New ADODB.Recordset
    
    If TAREAS_ Then
        ' Se llena las Tareas
        cSQL = "SELECT pro_cronogramatarea.idtar, pro_cronogramatarea.orden, pro_tareas.descripcion AS destar, pro_cronogramatarea.idcr, pro_cronogramatarea.idcrdet, pro_cronogramatarea.idlinea, pro_cronogramatarea.activo, pro_cronogramatarea.cantproc, pro_cronogramatarea.numper, pro_cronogramatarea.horinitar, pro_cronogramatarea.horfintar, pro_cronogramatarea.durtar, pro_cronogramatarea.fchini, pro_cronogramatarea.fchfin, pro_cronogramatarea.aplpor, pro_cronogramatarea.idresp, pla_empleados.nombre AS nomresp " _
            + vbCr + "FROM (pro_cronograma RIGHT JOIN ((pro_cronogramatarea LEFT JOIN pro_tareas ON pro_cronogramatarea.idtar = pro_tareas.id) LEFT JOIN alm_inventario ON pro_cronogramatarea.idpro = alm_inventario.id) ON pro_cronograma.id = pro_cronogramatarea.idcr) LEFT JOIN pla_empleados ON pro_cronogramatarea.idresp = pla_empleados.id " _
            + vbCr + "WHERE (((pro_cronograma.semana)=" & SEMANA_ & "));"
        
        RST_Busq xRs, cSQL, xCon
        If RstTareas.State = 0 Then
            DEFINIR_RST_TMP RstTareas, xRs
            DEFINIR_RST_TMP RstTareasAux, xRs
        Else
            limpiarRST RstTareas
            limpiarRST RstTareasAux
        End If
            
        CARGAR_RST_TMP RstTareas, xRs
        'CARGAR_RST_TMP RstTareasAux, xRs
        Set xRs = Nothing
    End If
    
    If PERSONAL_ Then
        ' Se llena al personal
        cSQL = "SELECT pro_cronogramapers.idper, pla_empleados.nombre, pla_empleados.codigo, pro_cronogramapers.idcr, pro_cronogramapers.idcrdet, pro_cronogramapers.idtar, pro_cronogramapers.activo, pro_tareas.descripcion AS destar " _
            + vbCr + "FROM pro_cronograma RIGHT JOIN ((pro_cronogramapers LEFT JOIN pla_empleados ON pro_cronogramapers.idper = pla_empleados.id) LEFT JOIN pro_tareas ON pro_cronogramapers.idtar = pro_tareas.id) ON pro_cronograma.id = pro_cronogramapers.idcr " _
            + vbCr + "WHERE (((pro_cronograma.semana)=" & SEMANA_ & "))"
        
        RST_Busq xRs, cSQL, xCon
        If RstPersonal.State = 0 Then
            DEFINIR_RST_TMP RstPersonal, xRs
            DEFINIR_RST_TMP RstPersonalAux, xRs
        Else
            limpiarRST RstPersonal
            limpiarRST RstPersonalAux
        End If
            
        CARGAR_RST_TMP RstPersonal, xRs
        'CARGAR_RST_TMP RstPersonalAux, xRs
        Set xRs = Nothing
    End If
    
    If RECETA_ Then
    End If
    
    If PRODUCTOS_ Then
        ' Se llena Productos
        cSQL = "SELECT pro_cronogramadet.*, alm_inventario.descripcion, mae_unidades.abrev, pro_receta.codrec " _
            + vbCr + "FROM (pro_cronograma LEFT JOIN ((pro_cronogramadet LEFT JOIN alm_inventario ON pro_cronogramadet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) ON pro_cronograma.id = pro_cronogramadet.idcr) LEFT JOIN pro_receta ON pro_cronogramadet.idrec = pro_receta.id " _
            + vbCr + "WHERE (((pro_cronograma.semana)=" & SEMANA_ & "))"
    
        RST_Busq xRs, cSQL, xCon
        
        If RstProductos.State = 0 Then
            DEFINIR_RST_TMP RstProductos, xRs
        Else
            limpiarRST RstProductos
        End If
            
        CARGAR_RST_TMP RstProductos, xRs
        Set xRs = Nothing
    End If
    
    If MATERIAPRIMA_ Then
        ' Se llena la materia Prima
        cSQL = "SELECT pro_cronogramadetprod.id, pro_cronogramadetprod.idcr, pro_cronogramadetprod.iditem, pro_cronogramadetprod.fchpro, pro_cronogramadetprod.horpro, pro_cronogramadetprod.idpro, pro_cronogramadetprod.cantidad, alm_inventario.descripcion AS descpro " _
            + vbCr + "FROM pro_cronograma RIGHT JOIN (pro_cronogramadetprod LEFT JOIN alm_inventario ON pro_cronogramadetprod.idpro = alm_inventario.id) ON pro_cronograma.id = pro_cronogramadetprod.idcr " _
            + vbCr + "WHERE (((pro_cronograma.semana)=" & SEMANA_ & ")) " _
            + vbCr + "GROUP BY pro_cronogramadetprod.id, pro_cronogramadetprod.idcr, pro_cronogramadetprod.iditem, pro_cronogramadetprod.fchpro, pro_cronogramadetprod.horpro, pro_cronogramadetprod.idpro, pro_cronogramadetprod.cantidad, alm_inventario.descripcion;"

        RST_Busq xRs, cSQL, xCon
        
        If RstMatPro.State = 0 Then
            DEFINIR_RST_TMP RstMatPro, xRs
        Else
            limpiarRST RstMatPro
        End If
            
        CARGAR_RST_TMP RstMatPro, xRs
        Set xRs = Nothing
    End If
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

Private Sub Form_Load()
    Agregando = False
    SeEjecuto = False
    QueHace = 3
    iniciarCampos
End Sub

Sub Modificar()
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Modificando Cronograma de Produccion"
    QueHace = 2
    xHorIni = Time
    ActivaTool
    Bloquea
    
    CmdOpciones(1).Enabled = True
    CmdOpciones(2).Enabled = True
    CmdOpciones(3).Enabled = True
    '******************************************************
    CmdOpciones(4).Enabled = True
    '******************************************************
    
    TxtIdSup.SetFocus
    
    ARRASTRANDO_ = False
End Sub

'*****************************************************************************************************
'* Nombre           : Nuevo
'* Tipo             : Sub
'* Descripcion      :
'* Modificacion     :
'*                    21/04/2011 JOSE CHACON
'*                      -> se modifica la referencia "id" de pro_cronogramadetprod por "idcr"
'*****************************************************************************************************
Sub Nuevo()
    Dim A As Integer
    
    QueHace = 1
    xHorIni = Time
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Label5.Caption = "Agregando Cronograma de Produccion"
    
    ActivaTool
    Blanquea
    Bloquea
    
    If RstProductos.State = 0 Then
        llenarDefinirRST 99999, True, False ' Tareas
        llenarDefinirRST 99999, False, True ' Personal
        llenarDefinirRST 99999, False, False, False, True ' Productos
        llenarDefinirRST 99999, False, False, False, False, True ' Materia Prima
    End If
    
    CmdBusTip.Enabled = True
    ComboSemanas.Locked = False
    CmdOpciones(0).Enabled = True
    
    CmdBusSup.SetFocus
    ARRASTRANDO_ = False
End Sub

Sub Bloquea()
    TxtIdSup.Locked = Not TxtIdSup.Locked
    TxtFchIni.Locked = Not TxtFchIni.Locked
    TxtFchFin.Locked = Not TxtFchFin.Locked
    TxtTipPro.Locked = Not TxtTipPro.Locked
    
    CmdBusSup.Enabled = Not CmdBusSup.Enabled
    
    habilitar Cmd, Not Cmd(0).Enabled
    OptPers(1).Value = True
    habilitar OptPers, Not OptPers(0).Enabled
    TxtOrden.Locked = False
    Cmd(3).Enabled = True
    Cmd(9).Enabled = True
    Cmd(11).Enabled = True
    'cmd(17).Enabled = True
    Cmd(19).Enabled = Not Cmd(19).Enabled
End Sub

Sub Blanquea()
    TxtIdSup.Text = ""
    TxtFchIni.valor = Date
    TxtTipPro.Text = ""
    LblTipoProd.Caption = ""
    LblSupervisor.Caption = ""
    CalendarControl1.DataProvider.RemoveAllEvents
    If QueHace = 1 Then CalendarControl1.Visible = False: CmdOpciones(0).Enabled = True
End Sub

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub

    If Me.Width <= 12000 Then Me.Width = 12000
    If Me.Height <= 8100 Then Me.Height = 8100

    ' Se dimensiona el Contenido
    TabOne1.Width = Me.Width - 150
    TabOne1.Height = Me.Height - 900

    Label6.Width = Me.Width - 100
    Dg1.Width = TabOne1.Width - 150
    Dg1.Height = TabOne1.Height - 850

    ' Se dimensiona el Detalle
    Label5.Width = Me.Width - 100

    CalendarControl1.Top = 1485
    CalendarControl1.Width = TabOne1.Width - 100
    CalendarControl1.Height = TabOne1.Height - 2450

    ShapeFondo.Width = CalendarControl1.Width
    ShapeFondo.Height = CalendarControl1.Height - 50

    FrmBotones.Top = TabOne1.Height - 1000
    FrmBotones.Width = TabOne1.Width - 100
End Sub

Sub Cancelar()
    QueHace = 3
    Bloquea
    Label5.Caption = "Consultando Cronograma de Produccion"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    CalendarControl1.Visible = True
    CmdOpciones(0).Enabled = False
    
    CmdBusTip.Enabled = False
    ComboSemanas.Locked = True
    CmdOpciones(0).Enabled = False
    
    CmdOpciones(1).Enabled = False
    CmdOpciones(2).Enabled = False
    CmdOpciones(3).Enabled = False
    
    '*****************************
    CmdOpciones(4).Enabled = False
    '*****************************
    ActivaTool
End Sub

Private Sub operaciones(Optional AGREGAR_ As Boolean = True, Optional MODIFICAR_ As Boolean = False, _
                                        Optional ELIMINAR_ As Boolean = False)
    If AGREGAR_ Then
        Dim FECHAINI_ As Date
        Dim FECHAFIN_ As Date
        Dim TODODIA_ As Boolean
        Dim EVENTONUEVO_ As CalendarEvent
    
        If QueHace <> 3 Then
            ' Se Elimina el evento anterior para reemplazarlo
            If Not EVENTO_ Is Nothing Then CalendarControl1.DataProvider.DeleteEvent EVENTO_
            ' Se crea el nuevo evento
            Set EVENTONUEVO_ = CalendarControl1.DataProvider.CreateEvent
            ' Se verfica el estado del recordset
            If RstProductos.State = 0 Then Exit Sub
            RstProductos.Filter = "id = " & NulosN(LblIdCrDet.Caption) & ""
            If RstProductos.RecordCount = 0 Then Exit Sub
            
            FECHAINI_ = CDate(Format(RstProductos("fchpro"), "dd/mm/yyyy") _
                                                    & " " + Format(RstProductos("horpro"), "HH:mm"))
            FECHAFIN_ = CDate(Format(RstProductos("fchfin"), "dd/mm/yyyy") _
                                                    & " " + Format(RstProductos("horfin"), "HH:mm"))
            ' Se llenan los datos
            EVENTONUEVO_.ScheduleID = NulosN(RstProductos("iditem"))                                        ' iditem
            EVENTONUEVO_.StartTime = FECHAINI_                                                              ' fech. Prog
            EVENTONUEVO_.EndTime = FECHAFIN_                                                                ' fech. Fin
            EVENTONUEVO_.Subject = NulosC(RstProductos("descripcion"))                                      ' Descripcion del Producto
            EVENTONUEVO_.Location = NulosN(RstProductos("numprod"))                                          ' Cantidad
            EVENTONUEVO_.Body = NulosN(RstProductos("cantidad")) & " " & NulosC(RstProductos("abrev")) _
                            + vbCr + Format(FECHAINI_, "HH:mm") & " - " & Format(FECHAFIN_, "HH:mm")        ' Hora de inicio y fin
            EVENTONUEVO_.ReminderSoundFile = NulosN(RstProductos("id"))                                     ' id de cronograma detallado
                                    
            CalendarControl1.DataProvider.AddEvent EVENTONUEVO_
            ' Se aumenta el correlativo
            CORR_ = CORR_ + 1
        End If
        
        FrmAdd.Visible = False
    End If
    
    If MODIFICAR_ Then
    End If
    
    If ELIMINAR_ Then
        Dim IDCRDET_ As Double
        Dim Rpta As Integer
        
        If DETECTOR_.ViewEvent Is Nothing Then Exit Sub
        ' Se encuentra el idcrdet involucrado
        IDCRDET_ = EVENTO_.ReminderSoundFile
        ' Se verifica el estado del recordset
        If RstProductos.State = 0 Then Exit Sub
        RstProductos.Filter = "id = " & IDCRDET_ & ""
        If RstProductos.RecordCount = 0 Then Exit Sub
        ' Se llena el recordset auxiliar
        If RstProductosAux.State = 0 Then DEFINIR_RST_TMP RstProductosAux, RstProductos
        limpiarRST RstProductosAux
        CARGAR_RST_TMP RstProductosAux, RstProductos
        limpiarRST RstProductos, False
        
        'xTitulo = "Eliminar evento"
        Rpta = MsgBox("¿Esta seguro de eliminar el evento seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
        If Rpta = vbYes Then
            ' Se elimina los recordsets relacionados
            RstTareas.Filter = "idcrdet = " & IDCRDET_ & ""
            RstPersonal.Filter = "idcrdet = " & IDCRDET_ & ""
            'RstReceta.Filter = "idcrdet = " & IDCRDET_ & ""
            limpiarRST RstTareas, False
            limpiarRST RstPersonal, False
            'limpiarRST RstReceta, False
            RstProductos.Filter = adFilterNone
            RstTareas.Filter = adFilterNone
            RstPersonal.Filter = adFilterNone
            'RstReceta.Filter = adFilterNone
            CalendarControl1.DataProvider.DeleteEvent EVENTO_
        Else
            CARGAR_RST_TMP RstProductos, RstProductosAux
            RstProductos.Filter = adFilterNone
        End If
            
        ' Se limpia el calendario
        CalendarControl1.DataProvider.RemoveAllEvents
        ' Se llenan todos los eventos
        LlenarDatos
    End If
    
    Set DETECTOR_ = Nothing
End Sub

Private Sub mostrarFormulario(Optional AGREGAR_ As Boolean = True, Optional MODIFICAR_ As Boolean = False, _
                                        Optional RECETA_ As Boolean = False)
    Dim fIni As Date
    Dim fFin As Date
    Dim AllDay As Boolean
        
    TabOne2.CurrTab = 0
    
    If AGREGAR_ Then ' Muestra el formulario pra ingreso de nuevos productos
        
        agregEvent = True
        modifEvent = False
        CalendarControl1.ActiveView.GetSelection fIni, fFin, AllDay
        
        ' SI es una fecha Incoherente
        If Format(fIni, "yyyy") < AnoTra Then Exit Sub
        
        LblDia.Caption = Format(fIni, "dd/mm/yyyy")         ' Detalle del Dia
        TxtNumProd.Text = ""                                ' Numero de Produccion
        LblIdCrDet.Caption = CORR_                          ' Correlativo
        TxtMatProd.Text = ""                                ' iditem
        LblMatProd.Caption = ""                             ' Descripcion prod
        '********************************************************************************
        TxtCodRec.Text = ""                                 ' codigo de receta
        lblIdRec.Caption = ""                               ' id receta
        '********************************************************************************
        TxtCant.Text = ""                                   ' Cantidad
        LblUnidad.Caption = ""                              ' UM
        DTPHoras.Value = Format(fIni, "HH:mm")              ' Hora de Inicio
        DTPHoraFin.Value = 0                                ' Hora de Fin
        TxtCantMP.Text = 0                                  ' Cantidad de MP
        TxtOrden.Text = ""                                  ' Orden de la tarea
        LblTarea.Caption = ""                               ' Detalle de la Tarea
        LblIdTarea.Caption = 0                              ' Id de la Tarea
        LblNTrab.Caption = 0                                ' Numero de trabajadores
        LblDetTrab.Caption = "0 de 0"                       ' Detalle de Trabajadores seleccionados
        TxtIdLineaDet.Text = ""                             ' id de Linea
        LblLinea.Caption = ""                               ' Detalle de Linea
        TxtIdEncarg.Text = ""
        LblEncargado.Caption = ""
        
        ' Se Agrega las Tareas
        pCargarDatos fg(0), False, True
        ' Se Agrega al personal
        pCargarDatos fg(1), True, False
        ' Se Agrega la Receta
        pCargarDatos fg(3), False, False, False, True
        
        OptPers(1).Value = True
        
        centrarFrm FrmAdd
        LblNomOperacion.Caption = "Agregando Cronograma"
        FrmAdd.Visible = True
    End If
    
    If MODIFICAR_ Then ' Muestra el formulario para modificar productos
        If DETECTOR_.ViewEvent Is Nothing Then Exit Sub
        modifEvent = True
        agregEvent = False
        
        IDCR_ = EVENTO_.ReminderSoundFile
        ' Buscamos el Producto involucrado
        RstProductos.Filter = "id = " & IDCR_ & ""
        ' Limpiamos el recordset Auxiliar
        limpiarRST RstProductosAux, True
        ' Cargamos el Recordset Auxiliar
        If RstProductosAux.State = 0 Then DEFINIR_RST_TMP RstProductosAux, RstProductos
        CARGAR_RST_TMP RstProductosAux, RstProductos
        
        ' Buscamos la Tarea involucrada
        RstTareas.Filter = "idcrdet = " & IDCR_ & ""
        ' Limpiamos el recordset Auxiliar
        limpiarRST RstTareasAux, True
        ' Cargamos el Recordset Auxiliar
        CARGAR_RST_TMP RstTareasAux, RstTareas
        
        LblDia.Caption = Format(RstProductosAux("fchpro"), "dd/mm/yyyy")        ' Dia de Programacion
        TxtNumProd.Text = NulosC(RstProductosAux("numprod"))
        LblIdCrDet.Caption = IDCR_                                              ' Correlativo
        TxtMatProd.Text = NulosN(RstProductosAux("iditem"))                     ' Id item
        LblMatProd.Caption = NulosC(RstProductosAux("descripcion"))             ' Descripcion Prod
        
        '*********************************************************************************************
        TxtCodRec.Text = NulosC(RstProductosAux("codrec"))                      ' Codigo de receta
        lblIdRec.Caption = NulosN(RstProductosAux("idrec"))                     ' Id receta
        '*********************************************************************************************
        
        TxtCant.Text = Format(NulosN(RstProductosAux("cantidad")), "0.00")      ' Cantidad
        LblUnidad.Caption = NulosC(RstProductosAux("abrev"))                    ' UM
        DTPHoras.Value = Format(RstProductosAux("horpro"), "HH:mm")             ' Hora de Inicio
        DTPHoraFin.Value = Format(RstProductosAux("horfin"), "HH:mm")           ' Hora de Fin
        TxtCantMP.Text = 0                                                      ' Cantidad de MP
        LblTarea.Caption = ""                                                   ' Detalle de la Tarea
        LblIdTarea.Caption = 0                                                  ' Id de la Tarea
        LblNTrab.Caption = 0                                                    ' Numero de trabajadores
        LblDetTrab.Caption = "0 de 0"                                           ' Detalle de Trabajadores seleccionados
        TxtOrden.Text = ""                                                       ' Orden de la tarea
        TxtIdEncarg.Text = ""
        LblEncargado.Caption = ""
        
        TxtOrden_Validate True
                
        ' Se Agrega las Tareas
        pCargarDatos fg(0), False, True
        ' Se Agrega al personal
        pCargarDatos fg(1), True, False
        ' Se Agrega la Receta
        pCargarDatos fg(3), False, False, False, True
        
        OptPers(1).Value = True
        
        centrarFrm FrmAdd
        LblNomOperacion.Caption = "Modificando Cronograma"
        FrmAdd.Visible = True
    End If
    
    If RECETA_ Then ' Muestra el formulario para escoger productos de la Materia prima
        If DETECTOR_.ViewEvent Is Nothing Then Exit Sub
        
        If QueHace = 3 Then
            CmdAcepta(0).Enabled = False
            Fg2.SelectionMode = flexSelectionByRow
            Fg2.Editable = flexEDNone
        Else
            CmdAcepta(0).Enabled = True
            Fg2.SelectionMode = flexSelectionFree
            Fg2.Editable = flexEDKbdMouse
        End If
        
        If TxtTipPro.Text <> 1 Then Exit Sub
        Dim Rst As New ADODB.Recordset
        Dim A, B As Integer
        Dim xMatPri As String
        
        Fg2.Rows = 1
        
        centrarFrm Frame3
        
        TxtMP.Text = EVENTO_.Subject
        TxtCan.Text = EVENTO_.ReminderSoundFile
        TxtCan.Text = Format(TxtCan.Text, "0.00")
        xMatPri = TxtMP.Text
            
        xFchPro = Mid(EVENTO_.StartTime, 1, 10)
        xHorPro = Mid(EVENTO_.StartTime, 11, 6)
        
        xIdMatPri = Busca_Codigo(xMatPri, "descripcion", "id", "alm_inventario", "C", xCon)
        
        If xIdMatPri = 0 Then
            MsgBox "La materia prima especificada no existe", vbInformation + vbOKOnly + vbDefaultButton1
            Exit Sub
        End If
        
        ' MOSTRAMOS TODOS LOS PRODUCTOS DE LA MATERIA PRIMA
        RST_Busq Rst, "SELECT pro_redimiento.iditem, pro_redimiento.idpro, alm_inventario.descripcion " _
            & " FROM pro_redimiento LEFT JOIN alm_inventario ON pro_redimiento.idpro = alm_inventario.id " _
            & " WHERE (((pro_redimiento.iditem)=" & xIdMatPri & "))", xCon
        
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            Fg2.Rows = 1
            For A = 1 To Rst.RecordCount
                Fg2.Rows = Fg2.Rows + 1
                Fg2.TextMatrix(A, 1) = Rst("descripcion")
                Fg2.TextMatrix(A, 2) = ""
                Fg2.TextMatrix(A, 3) = 0
                Fg2.TextMatrix(A, 4) = Rst("idpro")
                Rst.MoveNext
                If Rst.EOF = True Then Exit For
            Next A
            
            If Rst.RecordCount = 1 Then
                Fg2.TextMatrix(Fg2.Rows - 1, 2) = Format(TxtCan.Text, "0.00")
                If QueHace = 3 Then
                    Fg2.TextMatrix(Fg2.Rows - 1, 3) = 0
                    Fg2.TextMatrix(Fg2.Rows - 1, 2) = ""
                Else
                    Fg2.TextMatrix(Fg2.Rows - 1, 3) = 1
                End If
                Fg2.Editable = flexEDNone
            End If
        End If
        
        ' MOSTRAMOS EL CHECK DE LOS PRODUCTOS QUE SE VAYAN A DEFINIR
        RstMatPro.Filter = adFilterNone
        If RstMatPro.RecordCount <> 0 Then
            RstMatPro.Filter = "iditem =" & xIdMatPri & " AND fchpro = " & xFchPro & " AND horpro = " & Format(xHorPro, "HH:mm") & ""
            If RstMatPro.RecordCount <> 0 Then
                RstMatPro.MoveFirst
                For A = 1 To RstMatPro.RecordCount
                    For B = 1 To Fg2.Rows - 1
                        If NulosN(Fg2.TextMatrix(B, 4)) = RstMatPro("idpro") Then
                            Fg2.TextMatrix(B, 3) = 1
                            Fg2.TextMatrix(B, 2) = Format(RstMatPro("cantidad"), "0.00")
                            Exit For
                        End If
                    Next B
                    RstMatPro.MoveNext
                    If RstMatPro.EOF = True Then Exit For
                Next A
            End If
        End If
        TxtTotal.Text = Format(GRID_SUMAR_COL(Fg2, 2, 1, Fg2.Rows - 1), "0.00")
        
        
        If visEvent Then CmdAcepta(0).Enabled = False
        Frame3.Visible = True
    End If
    TxtNumProd.SetFocus
End Sub

Private Sub centrarFrm(ByRef frm As Frame)
    With frm
        .Left = ((Me.Width - .Width) / 2)
        .Top = ((Me.Height - .Height) / 2)
    End With
End Sub

Private Function encontrarUnidad(idProd As String) As String
    Dim codigo As String
    Dim unidad As String
    codigo = Busca_Codigo(idProd, "id", "idunimed", "alm_inventario", "N", xCon)
    If NulosC(codigo) <> "" Then
        unidad = Busca_Codigo(codigo, "id", "abrev", "mae_unidades", "N", xCon)
    Else
        unidad = ""
    End If
    encontrarUnidad = unidad
End Function


Private Sub menu_01_01_Click()
    ' Productos de receta
    mostrarFormulario False, False, True
End Sub

Private Sub menu_01_02_Click()
    ' Agregar
    mostrarFormulario
End Sub

Private Sub menu_01_03_Click()
    ' Eliminar
    operaciones False, False, True
End Sub

Private Sub menu_01_04_Click()
    ' Modificar
    mostrarFormulario False, True, False
End Sub

Private Sub Menu2_1_Click()
    ' Agregar
    mostrarFormulario
End Sub

Private Sub menu2_2_Click()
    ' Eliminar
    operaciones False, False, True
End Sub

Private Sub Menu2_3_Click()
    ' Modificar
    mostrarFormulario False, True, False
End Sub

Private Sub Menu3_1_Click()
    ' Productos de receta
    mostrarFormulario False, False, True
End Sub

Private Sub OptHoras_Click(Index As Integer)
    If Index = 0 Then
        DTPHorIni.Enabled = True
        DTPHorFin.Enabled = True
    End If
    
    If Index = 1 Then
        DTPHorIni.Enabled = False
        DTPHorFin.Enabled = False
    End If
End Sub

Private Sub OptPers_Click(Index As Integer)
    If Index = 0 Then
        TxtOrden.Text = ""
        TxtOrden.Locked = True
        LblTarea.Caption = ""
        Cmd(3).Enabled = False
        fg(1).ColWidth(3) = 3100
        fg(1).ColWidth(7) = 2500
        pCargarDatos fg(1), True, False, True
    End If
    If Index = 1 Then
        TxtOrden.Locked = False
        Cmd(3).Enabled = True
        fg(1).ColWidth(3) = 3810
        fg(1).ColWidth(7) = 0
        TxtOrden_Validate True
    End If
End Sub

Private Sub optTarea_Click(Index As Integer)
    If Index = 0 Then
        TxtPctje.Enabled = False
        DTPMinutos.Enabled = False
    End If
    
    If Index = 1 Then
        TxtPctje.Enabled = True
        TxtPctje.SetFocus
        DTPMinutos.Enabled = False
    End If
    
    If Index = 2 Then
        DTPMinutos.Enabled = True
        TxtPctje.Enabled = False
    End If
    
    If Index = 3 Then
    End If
End Sub

Private Sub PbCerrar_Click(Index As Integer)
    If Index = 0 Then
        FrmAdd.Visible = False
    End If
    
    If Index = 1 Then
        TabOne1.Enabled = True
        Toolbar1.Enabled = True
        
        CmdAcepta(0).Enabled = True
        visEvent = False
        Frame3.Visible = False
    End If
    
    If Index = 2 Then
        Frame9.Visible = False
    End If
    
    If Index = 3 Then
        Frame10.Visible = False
    End If
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    ' Se limpia el calendario
    CalendarControl1.DataProvider.RemoveAllEvents
    If OldTab = 0 Then
        If QueHace = 1 Then Exit Sub
        MuestraSegundoTab
    Else
        Frame3.Visible = False
        FrmAdd.Visible = False
        Frame9.Visible = False
        Frame10.Visible = False
        CARGO_ = False
        limpiarRST RstProductos
        Set RstProductos = Nothing
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : Sub
'* Descripcion      :
'* Modificacion     :
'*                    21/04/2011 JOSE CHACON
'*                      -> se modifica la referencia "id" de pro_cronogramadet, pro_cronogramadetprod por "idcr"
'*                    03/05/2011 JOSE CHACON
'*                      -> Se agrega la eliminacion de la tabla pro_cronogramapers
'*****************************************************************************************************
Sub Eliminar()
    TabOne1.CurrTab = 0
    Dim Rpta As Integer
    Dim idregistro As Double
    Dim xId As Double
    Dim xRs As New ADODB.Recordset
    
    Rpta = MsgBox("¿Esta seguro de eliminar el cronograma seleccionado?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
    If Rpta = vbYes Then
        'Busco todos los cronogramas relacionados con esa semana
        cSQL = "SELECT pro_cronograma.id AS idcr, pro_cronograma.semana " _
            + vbCr + "From pro_cronograma " _
            + vbCr + "Where (((pro_cronograma.semana) = " & NulosN(RstLis("semana")) & ")) " _
            + vbCr + "GROUP BY pro_cronograma.id, pro_cronograma.semana;"
        
        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        idregistro = NulosN(RstLis("semana"))
        xRs.MoveFirst
        While Not xRs.EOF
            xId = NulosN(xRs("idcr"))
            xCon.Execute "DELETE * FROM pro_cronogramapers WHERE idcr = " & xId & ""
            xCon.Execute "DELETE * FROM pro_cronogramatarea WHERE idcr = " & xId & ""
            xCon.Execute "DELETE * FROM pro_cronogramadetprod WHERE idcr = " & xId & ""
            xCon.Execute "DELETE * FROM pro_cronogramadet WHERE idcr = " & xId & ""
            xCon.Execute "DELETE * FROM pro_cronograma WHERE id = " & xId & ""
            
            xRs.MoveNext
        Wend
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & idregistro & " AND idform = " & IdMenuActivo
        
        RstLis.Requery
        Dg1.Refresh
        'xTitulo = "Grabar"
        MsgBox "El cronograma se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstLis.Requery
            Dg1.Refresh
            If RstLis.RecordCount <> 0 Then
                RstLis.MoveFirst
                RstLis.Find "semana=" & mIdRegistro
                If RstLis.EOF = True Then RstLis.MoveFirst
            End If
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 9 Then
        If TabOne1.CurrTab = 0 Then RstLis.Filter = "": TDB_FiltroLimpiar Dg1
        If TabOne1.CurrTab = 1 Then CmdOpciones_Click 0
    End If
    
    If Button.Index = 12 Then
'        If TabOne1.CurrTab = 0 Then Exit Sub
'        IMPRIMIR
    End If
    
    If Button.Index = 14 Then
        Set RstLis = Nothing
        Unload Me
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then ' Imprimir linea
        imprimir 0
    End If
    If ButtonMenu.Index = 2 Then ' Imprimir Acabado
    End If
    If ButtonMenu.Index = 3 Then ' Imprimir Reporte
    End If
End Sub

Private Sub exportarPDF()
    Dim Rst As New ADODB.Recordset

    If NulosN(RstLis("idtippro")) = 1 Then
        RST_Busq Rst, "TRANSFORM sum(pro_cronogramadetprod.cantidad) AS PromedioDecantidad SELECT pro_cronogramadetprod.iditem, alm_inventario_1.descripcion AS desmatpri, " _
            & " mae_unidades.abrev, alm_inventario.descripcion AS descprod, Sum(pro_cronogramadetprod.cantidad) AS [TotalFila]" _
            & " FROM ((pro_cronogramadetprod LEFT JOIN alm_inventario ON pro_cronogramadetprod.idpro = alm_inventario.id) LEFT JOIN alm_inventario AS alm_inventario_1 " _
            & " ON pro_cronogramadetprod.iditem = alm_inventario_1.id) LEFT JOIN mae_unidades ON alm_inventario_1.idunimed = mae_unidades.id " _
            & " Where (((pro_cronogramadetprod.ID) = " & RstLis("id") & ")) GROUP BY pro_cronogramadetprod.iditem, alm_inventario_1.descripcion, mae_unidades.abrev, " _
            & " alm_inventario.descripcion, pro_cronogramadetprod.id ORDER BY alm_inventario_1.descripcion, alm_inventario.descripcion " _
            & " PIVOT Format([fchpro],'dd-mm-yy')", xCon
    Else
        RST_Busq Rst, "TRANSFORM Sum(pro_cronogramadet.cantidad) AS SumaDecantidad SELECT pro_cronogramadet.iditem, alm_inventario.descripcion, mae_unidades.abrev, " _
            & " Sum(pro_cronogramadet.cantidad) AS TotalFila FROM (pro_cronogramadet LEFT JOIN alm_inventario ON pro_cronogramadet.iditem = alm_inventario.id) " _
            & " LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id Where (((pro_cronogramadet.ID) = " & RstLis("id") & ")) " _
            & " GROUP BY pro_cronogramadet.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_cronogramadet.id PIVOT Format([fchpro],'dd-mm-yy')", xCon
    End If

    Dim Li As Integer
    Dim strSource As String
    Dim xArea, xEmp, xDir, xCuerpo, xCad  As String
    Dim xEmpleado As String
    Dim Pagina As Integer
    Dim Lineas As Integer

    Set oPDF = New cPDF
    Dim A, B, C As Integer
    xNumPag = 0
    Dim xTipPro As String

On Error GoTo Cerrado

    If oPDF.PDFCreate(App.Path & "\pro00001.pdf") = True Then

        oPDF.Fonts.Add "Tit", Times_BoldItalic, WinAnsiEncoding
        oPDF.Fonts.Add "Head", Times_Italic, WinAnsiEncoding
        oPDF.Fonts.Add "Cont", Courier, WinAnsiEncoding
        oPDF.Fonts.Add "CB", Courier_Bold, WinAnsiEncoding
        oPDF.Fonts.Add "Time", Times_Roman, WinAnsiEncoding

        CrearCabecera
        Dim xFilaAct As Integer
        Dim xPosX As Integer
        Dim xFch As Date

        oPDF.WTextBox 40, 30, 10, 750, "CRONOGRAMA DE PRODUCCION (" & RstLis("destippro") & ")", "CB", 10, hCenter, vMiddle, vbBlack, 0, vbRed
        oPDF.WTextBox 52, 30, 10, 750, "DEL " & RstLis("fchini") & " AL " & RstLis("fchfin"), "CB", 10, hCenter, vMiddle, vbBlack, 0, vbRed

        If NulosN(RstLis("idtippro")) = 1 Then
            oPDF.WTextBox 68, 30, 19, 150, "MATERIA PRIMA", "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack
            oPDF.WTextBox 68, 180, 19, 30, "UNI. MED.", "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack
            oPDF.WTextBox 68, 210, 19, 250, "PRODUCTO", "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack
            xPosX = 460
        Else
            oPDF.WTextBox 68, 30, 19, 250, "PRODUCTO", "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack
            oPDF.WTextBox 68, 280, 19, 30, "UNI. MED.", "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack
            xPosX = 310
        End If

        ' IMPRIMIMOS EL ROTULO DE LAS FECHAS
        For xFch = RstLis("fchini") To RstLis("fchfin")
            oPDF.WTextBox 68, xPosX, 19, 45, Format(xFch, "dd/mm/yy"), "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack

            xPosX = xPosX + 45
        Next xFch

        ' IMPRIMIMOS EL ROTULO DEL TOTAL
        oPDF.WTextBox 68, xPosX, 19, 45, "TOTAL", "CB", 8, hCenter, vMiddle, vbBlack, 1, vbBlack

        xFilaInicial = 88
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            For A = 1 To Rst.RecordCount
                If NulosN(RstLis("idtippro")) = 1 Then
                    oPDF.WTextBox xFilaInicial, 30, 10, 150, Rst("desmatpri"), "CB", 8, hLeft, vMiddle, vbBlack, 0, vbBlack
                    oPDF.WTextBox xFilaInicial, 180, 10, 30, Rst("abrev"), "CB", 8, hCenter, vMiddle, vbBlack, 0, vbRed
                    oPDF.WTextBox xFilaInicial, 210, 10, 250, Rst("descprod"), "CB", 8, hLeft, vMiddle, vbBlack, 0, vbRed
                    xPosX = 460
                Else
                    oPDF.WTextBox xFilaInicial, 30, 10, 250, Rst("descripcion"), "CB", 8, hLeft, vMiddle, vbBlack, 0, vbBlack
                    oPDF.WTextBox xFilaInicial, 280, 10, 30, Rst("abrev"), "CB", 8, hCenter, vMiddle, vbBlack, 0, vbRed
                    xPosX = 310
                End If

                For xFch = RstLis("fchini") To RstLis("fchfin")
                    If RstRegistroBuscaCampo(Rst, Format(xFch, "dd-mm-yy")) = True Then
                        oPDF.WTextBox xFilaInicial, xPosX, 10, 45, Format(NulosN(Rst(Format(xFch, "dd-mm-yy"))), "0.00"), "CB", 8, hRight, vMiddle, vbBlack, 0, vbBlack
                    End If
                    xPosX = xPosX + 45
                Next xFch

                ' IMPRIMIMOS EL TOTAL DE LA FILA
                oPDF.WTextBox xFilaInicial, xPosX, 10, 45, Format(NulosN(Rst("TotalFila")), "0.00"), "CB", 8, hRight, vMiddle, vbBlack, 0, vbBlack
                xFilaInicial = xFilaInicial + 10

                Rst.MoveNext
                If Rst.EOF = True Then Exit For
            Next A
        End If

        oPDF.PDFClose
        Set oPDF = Nothing
        Shell ("rundll32.exe url.dll,FileProtocolHandler " & Trim(App.Path) & ("\pro00001.pdf")), vbMaximizedFocus
    Else
        Set oPDF = Nothing
        MsgBox "No se Puede Mostrar Documento pro00001.pdf, psoblemente el archivo ya se encuentra abierto", vbCritical, "Error"
    End If
    Exit Sub

Cerrado:
    'Resume
    If Err.Number = 1 Then
    End If
End Sub

Sub CrearCabecera()
    Dim xTelEmp, xNumDoc As String
    
    'oPDF.NewPage A4_Vertical ', 525, 675
    oPDF.NewPage A4_Horizontal  ', 525, 675
    xNumPag = xNumPag + 1
    
    oPDF.WTextBox 15, 30, 8, 50, "EMPRESA", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 15, 105, 8, 5, ":", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 15, 111, 8, 150, NomEmp, "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    
    oPDF.WTextBox 23, 30, 8, 50, "Nº R.U.C.", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 23, 105, 8, 5, ":", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 23, 111, 8, 100, NumRUC, "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    
    oPDF.WTextBox 15, 700, 8, 50, "Nº PAGINA", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 15, 750, 8, 5, ":", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 15, 753, 8, 50, Format(xNumPag, "000"), "CB", 8, hRight, vTop, RGB(0, 0, 128), , vbRed
    
    oPDF.WTextBox 23, 700, 8, 50, "FCH. IMPR", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 23, 750, 8, 5, ":", "CB", 8, hLeft, vTop, RGB(0, 0, 128), , vbRed
    oPDF.WTextBox 23, 753, 8, 50, Format(Date, "dd/mm/yy"), "CB", 8, hRight, vTop, RGB(0, 0, 128), , vbRed
    
    oPDF.WLineTo 30, 36, 800, 36
    oPDF.LineStroke
End Sub

Private Sub TxtCant_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtCant_Validate(Cancel As Boolean)
    TxtCant.Text = Format(NulosN(TxtCant.Text), "0.00")
End Sub

Private Sub TxtCantMP_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub TxtCodRec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtFchFin_Change()
    Dim fech As String
    If Not CAMBIO_ Then
        If TxtFchFin.valor <> "" Then
            fech = TxtFchFin.valor
            ComboSemanas.Text = DatePart("ww", NulosC(CDate(fech)), vbMonday, vbFirstFullWeek)
        End If
    End If
End Sub

Private Sub TxtFchIni_Change()
    Dim fech As String
    If Not CAMBIO_ Then
        If TxtFchIni.valor <> "" Then
            fech = TxtFchIni.valor
            ComboSemanas.Text = DatePart("ww", NulosC(CDate(fech)), vbMonday, vbFirstFullWeek)
        End If
    End If
End Sub

Private Sub TxtIdEncarg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdEncarg_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Cmd_Click 18
    End If
End Sub

Private Sub TxtIdLineaDet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdLineaDet_Validate(Cancel As Boolean)
    If NulosN(TxtIdLineaDet.Text) = 0 Then
        TxtIdLineaDet.Text = ""
        LblLinea.Caption = ""
        Exit Sub
    Else
        Dim codigo As String
        LblLinea.Caption = Busca_Codigo(TxtIdLineaDet.Text, "id", "descripcion", "pro_linea", "N", xCon)
    End If
End Sub

Private Sub TxtIdSup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtIdSup_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusSup_Click
    End If
End Sub

Private Sub TxtIdSup_Validate(Cancel As Boolean)
    If NulosN(TxtIdSup.Text) = 0 Then
        TxtIdSup.Text = ""
        Exit Sub
    Else
        Dim Rst As New ADODB.Recordset
        Dim xSqlCad As String
        xSqlCad = "SELECT pro_emp.*, pla_empleados.nombre, pro_emp.id " _
            & " FROM (pro_emp LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id " _
            & " Where (((pro_empdet.idfun) = 2) And ((pro_emp.ID) = " & Val(TxtIdSup.Text) & ")) ORDER BY pla_empleados.nombre"

        Set Rst = BuscaConCriterio(xSqlCad, xCon)
        
        If Rst.RecordCount <> 0 Then
            LblSupervisor.Caption = Rst("nombre")
        Else
            TxtIdSup.Text = ""
            LblSupervisor.Caption = ""
        End If
        
        Set Rst = Nothing
    End If
End Sub

Private Sub TxtMatProd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtMatProd_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        Cmd_Click 0
    End If
End Sub

Private Sub TxtMatProd_Validate(Cancel As Boolean)
    If NulosN(TxtMatProd.Text) = 0 Then
        TxtMatProd.Text = ""
        Exit Sub
    Else
        Dim codigo As String
        LblMatProd.Caption = Busca_Codigo(TxtMatProd.Text, "id", "descripcion", "alm_inventario", "N", xCon)
        codigo = Busca_Codigo(TxtMatProd.Text, "id", "idunimed", "alm_inventario", "N", xCon)
        If NulosC(codigo) <> "" Then LblUnidad.Caption = Busca_Codigo(codigo, "id", "abrev", "mae_unidades", "N", xCon)
        If NulosC(LblMatProd.Caption) = "" Then
            TxtMatProd.Text = ""
            LblUnidad.Caption = ""
            TxtMatProd.SetFocus
        Else
            TxtCant.SetFocus
        End If
    End If
End Sub

Private Sub TxtNumProd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtOrden_Validate(Cancel As Boolean)
    Dim RstAux As New ADODB.Recordset
    
    If RstLis.RecordCount = 0 Then Exit Sub
    
    cSQL = "SELECT pro_tareas.id, pro_tareas.descripcion, pro_cronogramatarea.orden, pro_cronogramatarea.numper " _
        + vbCr + "FROM pro_cronograma LEFT JOIN (pro_cronogramatarea LEFT JOIN pro_tareas ON pro_cronogramatarea.idtar = pro_tareas.id) ON pro_cronograma.id = pro_cronogramatarea.idcr " _
        + vbCr + "WHERE (((pro_cronograma.semana)= " & NulosN(RstLis("semana")) & ") AND ((pro_cronogramatarea.idpro)= " & NulosN(TxtMatProd.Text) & ") AND ((pro_cronogramatarea.orden)= " & NulosN(TxtOrden.Text) & "));"
        
    RST_Busq RstAux, cSQL, xCon
    
    If Not RstAux.EOF Then
        LblTarea.Caption = NulosC(RstAux("descripcion"))
        LblIdTarea.Caption = NulosN(RstAux("id"))
        LblNTrab.Caption = NulosN(RstAux("numper"))
        pCargarDatos fg(1), True, False
    Else
        LblTarea.Caption = ""
        LblIdTarea.Caption = ""
        LblNTrab.Caption = ""
        pCargarDatos fg(1), True, False
    End If
End Sub

Private Sub TxtPctje_Change()
    TxtPctje.Text = Format(NulosN(TxtPctje.Text), "0.00")
End Sub

Private Sub TxtPctje_GotFocus()
    Me.TxtPctje.SelStart = 0
    Me.TxtPctje.SelLength = Len(Me.TxtPctje.Text)
End Sub

Private Sub TxtTipPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtTipPro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTip_Click
    End If
End Sub

Private Sub TxtTipPro_Validate(Cancel As Boolean)
    If NulosN(TxtTipPro.Text) = 0 Then
        TxtTipPro.Text = ""
        Exit Sub
    Else
        LblTipoProd.Caption = Busca_Codigo(TxtTipPro.Text, "id", "descripcion", " mae_tipoproducto", "N", xCon)
        If NulosC(LblTipoProd.Caption) = "" Then
            TxtTipPro.Text = ""
        End If
    End If
End Sub

'Metodos para arrastrar el Frame
''''''''''''''''''''''''''''''''
Private Sub FrmAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    FrmAdd.ZOrder 0
End Sub

Private Sub FrmAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With FrmAdd
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub

Private Sub Frame3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    Frame3.ZOrder 0
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With Frame3
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub

Private Sub Frame9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    Frame9.ZOrder 0
End Sub

Private Sub Frame9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With Frame9
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub

Private Sub Frame10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    Frame10.ZOrder 0
End Sub

Private Sub Frame10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With Frame10
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub
