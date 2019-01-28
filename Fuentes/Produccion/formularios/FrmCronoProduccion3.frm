VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCronoProduccion3 
   Caption         =   "Producción - Programación Semanal"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   11820
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Height          =   1680
      Index           =   1
      Left            =   9450
      TabIndex        =   52
      Top             =   8790
      Visible         =   0   'False
      Width           =   5610
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   5310
         Picture         =   "FrmCronoProduccion3.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   57
         ToolTipText     =   "Cerrar"
         Top             =   60
         Width           =   195
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "Aceptar"
         Height          =   330
         Index           =   17
         Left            =   2820
         TabIndex        =   56
         ToolTipText     =   "Aceptar Seleccion"
         Top             =   1260
         Width           =   1300
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "&Cancelar"
         Height          =   330
         Index           =   9
         Left            =   4200
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "Cancelar Seleccion"
         Top             =   1260
         Width           =   1300
      End
      Begin VB.Frame Frame9 
         Caption         =   "[ Migrar a ]"
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
         Height          =   555
         Left            =   60
         TabIndex        =   53
         Top             =   690
         Width           =   5445
         Begin VB.ComboBox cbfchcamb 
            Height          =   315
            Left            =   3870
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   180
            Width           =   1365
         End
         Begin VB.ComboBox cbsemcamb 
            Height          =   315
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   180
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Semana"
            Height          =   195
            Left            =   1590
            TabIndex        =   60
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Left            =   3330
            TabIndex        =   54
            Top             =   210
            Width           =   450
         End
      End
      Begin VB.Label LblDetProd 
         AutoSize        =   -1  'True
         Caption         =   "difdia"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   2130
         TabIndex        =   65
         Top             =   1350
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label LblDetProd 
         AutoSize        =   -1  'True
         Caption         =   "LblIdCrDetDest"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   930
         TabIndex        =   64
         Top             =   1350
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Migrar Evento"
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
         Left            =   45
         TabIndex        =   59
         Top             =   75
         Width           =   1200
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
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   0
         X1              =   5580
         X2              =   5580
         Y1              =   0
         Y2              =   1650
      End
      Begin VB.Label LblDetProd 
         AutoSize        =   -1  'True
         Caption         =   "LblIdCrDet"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   58
         Top             =   1350
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   5580
         Y1              =   1650
         Y2              =   1650
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   7
         Left            =   10
         Top             =   45
         Width           =   5530
      End
      Begin VB.Label LblProd 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblProd"
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   60
         TabIndex        =   63
         Top             =   360
         Width           =   5445
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
      Left            =   1500
      TabIndex        =   37
      Top             =   7680
      Visible         =   0   'False
      Width           =   7530
      Begin VB.Frame Frame6 
         ForeColor       =   &H00800000&
         Height          =   3915
         Left            =   60
         TabIndex        =   75
         Top             =   300
         Width           =   7365
         Begin VB.CommandButton Cmd 
            Caption         =   "&Cancelar"
            Height          =   350
            Index           =   15
            Left            =   5910
            TabIndex        =   80
            ToolTipText     =   "Cancelar Edicion del Producto"
            Top             =   3450
            Width           =   1400
         End
         Begin VB.CommandButton Cmd 
            Caption         =   "&Aceptar"
            Height          =   330
            Index           =   14
            Left            =   5910
            TabIndex        =   79
            ToolTipText     =   "Eliminar Todos"
            Top             =   3060
            Width           =   1400
         End
         Begin VB.CommandButton Cmd 
            Caption         =   "&Procesar"
            Enabled         =   0   'False
            Height          =   350
            Index           =   2
            Left            =   5910
            TabIndex        =   77
            ToolTipText     =   "Procesar las Tareas del Producto Seleccionado"
            Top             =   240
            Width           =   1400
         End
         Begin VB.CommandButton Cmd 
            Caption         =   "&Propiedades"
            Enabled         =   0   'False
            Height          =   350
            Index           =   1
            Left            =   5910
            TabIndex        =   76
            ToolTipText     =   "Mostrar Propiedades de Procesado de Tareas"
            Top             =   630
            Width           =   1400
         End
         Begin VSFlex7Ctl.VSFlexGrid fg 
            Height          =   3570
            Index           =   0
            Left            =   90
            TabIndex        =   78
            Top             =   240
            Width           =   5780
            _cx             =   10195
            _cy             =   6297
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
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmCronoProduccion3.frx":02EC
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
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   3
         Left            =   7260
         Picture         =   "FrmCronoProduccion3.frx":0362
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   38
         ToolTipText     =   "Cerrar"
         Top             =   30
         Width           =   195
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
         Caption         =   "Detalle de Tareas"
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
         TabIndex        =   39
         Top             =   60
         Width           =   1530
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
   Begin VB.Frame frm 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Index           =   3
      Left            =   90
      TabIndex        =   40
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
         Top             =   480
         Width           =   1770
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
      Height          =   5040
      Index           =   0
      Left            =   12000
      TabIndex        =   11
      Top             =   390
      Visible         =   0   'False
      Width           =   4980
      Begin VB.Frame Frame3 
         Caption         =   "Opciones Diversas"
         Height          =   1125
         Left            =   150
         TabIndex        =   48
         Top             =   3360
         Width           =   4635
         Begin VB.CheckBox ckperarea 
            Caption         =   "Limitar Seleccion de Personal por Area"
            Height          =   195
            Left            =   180
            TabIndex        =   51
            Top             =   810
            Width           =   3045
         End
         Begin VB.CheckBox cknumper 
            Caption         =   "Limitar Numero de Personal segun Linea"
            Height          =   195
            Left            =   180
            TabIndex        =   50
            Top             =   540
            Width           =   3285
         End
         Begin VB.CheckBox cknumtar 
            Caption         =   "Limitar Numero de Tareas segun Linea"
            Height          =   195
            Left            =   180
            TabIndex        =   49
            Top             =   270
            Width           =   3195
         End
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "Aceptar"
         Height          =   345
         Index           =   12
         Left            =   2430
         TabIndex        =   33
         Top             =   4570
         Width           =   1155
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "&Cancelar"
         Height          =   345
         Index           =   13
         Left            =   3645
         TabIndex        =   32
         Top             =   4570
         Width           =   1155
      End
      Begin VB.Frame Frame11 
         Caption         =   "Incluir Horas de Refrigerio?"
         Height          =   945
         Left            =   150
         TabIndex        =   23
         Top             =   2400
         Width           =   4660
         Begin VB.OptionButton OptHoras 
            Caption         =   "No"
            Height          =   225
            Index           =   1
            Left            =   1000
            TabIndex        =   25
            Top             =   450
            Width           =   615
         End
         Begin VB.OptionButton OptHoras 
            Caption         =   "Si"
            Height          =   225
            Index           =   0
            Left            =   300
            TabIndex        =   24
            Top             =   450
            Width           =   555
         End
         Begin MSComCtl2.DTPicker DTPHorIni 
            Height          =   345
            Left            =   2700
            TabIndex        =   26
            Top             =   130
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "HH:mm"
            Format          =   16973827
            UpDown          =   -1  'True
            CurrentDate     =   40606
         End
         Begin MSComCtl2.DTPicker DTPHorFin 
            Height          =   345
            Left            =   2700
            TabIndex        =   27
            Top             =   500
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "HH:mm"
            Format          =   16973827
            UpDown          =   -1  'True
            CurrentDate     =   40606
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "( Fin"
            Height          =   195
            Index           =   9
            Left            =   2100
            TabIndex        =   31
            Top             =   585
            Width           =   300
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "HH:mm )"
            Height          =   195
            Index           =   10
            Left            =   3705
            TabIndex        =   30
            Top             =   585
            Width           =   615
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "( Inicio"
            Height          =   195
            Index           =   30
            Left            =   2100
            TabIndex        =   29
            Top             =   225
            Width           =   465
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "HH:mm )"
            Height          =   195
            Index           =   8
            Left            =   3700
            TabIndex        =   28
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
         Picture         =   "FrmCronoProduccion3.frx":064E
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   22
         ToolTipText     =   "Cerrar"
         Top             =   60
         Width           =   195
      End
      Begin VB.Frame Frame12 
         Caption         =   "La tarea Empieza al : "
         Height          =   2085
         Left            =   150
         TabIndex        =   12
         Top             =   300
         Width           =   4660
         Begin VB.OptionButton OptTarea 
            Caption         =   "Finalizar la tarea anterior"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   17
            Top             =   270
            Width           =   2775
         End
         Begin VB.OptionButton OptTarea 
            Caption         =   "Transcurrir un porcentaje de la tarea anterior"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   16
            Top             =   510
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
            TabIndex        =   15
            Text            =   "TxtPctje"
            Top             =   795
            Width           =   840
         End
         Begin VB.OptionButton OptTarea 
            Caption         =   "Transcurrido los minutos de la tarea anterior"
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   14
            Top             =   1140
            Width           =   3855
         End
         Begin VB.OptionButton OptTarea 
            Caption         =   "Segun Linea"
            Height          =   255
            Index           =   3
            Left            =   210
            TabIndex        =   13
            Top             =   1770
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker DTPMinutos 
            Height          =   345
            Left            =   2160
            TabIndex        =   35
            Top             =   1410
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "HH:mm"
            Format          =   16973827
            UpDown          =   -1  'True
            CurrentDate     =   40606
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Porcentaje"
            Height          =   195
            Index           =   6
            Left            =   1245
            TabIndex        =   21
            Top             =   840
            Width           =   765
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Index           =   2
            Left            =   3075
            TabIndex        =   20
            Top             =   840
            Width           =   120
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Minutos"
            Height          =   195
            Index           =   7
            Left            =   1245
            TabIndex        =   19
            Top             =   1440
            Width           =   555
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "HH:mm"
            Height          =   195
            Index           =   4
            Left            =   3075
            TabIndex        =   18
            Top             =   1440
            Width           =   525
         End
      End
      Begin VB.Label lbl 
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
         Index           =   27
         Left            =   105
         TabIndex        =   34
         Top             =   50
         Width           =   2955
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   2
         X1              =   0
         X2              =   4950
         Y1              =   5000
         Y2              =   5000
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   2
         X1              =   4950
         X2              =   4950
         Y1              =   0
         Y2              =   5000
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
   Begin VB.Frame frm 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Index           =   5
      Left            =   5370
      TabIndex        =   44
      Top             =   7620
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   225
         TabIndex        =   45
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
         Height          =   195
         Index           =   32
         Left            =   225
         TabIndex        =   47
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
         TabIndex        =   46
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
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7125
      Left            =   30
      TabIndex        =   1
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
         TabIndex        =   3
         Top             =   390
         Width           =   11760
         Begin VB.Frame FrmBotones 
            Caption         =   "[ Programación]"
            Height          =   6330
            Left            =   0
            TabIndex        =   7
            Top             =   300
            Width           =   11750
            Begin VB.ComboBox cbSemana 
               Height          =   315
               ItemData        =   "FrmCronoProduccion3.frx":093A
               Left            =   840
               List            =   "FrmCronoProduccion3.frx":093C
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   73
               Top             =   270
               Width           =   1155
            End
            Begin VB.ComboBox cbfecha 
               Height          =   315
               Left            =   10410
               Style           =   2  'Dropdown List
               TabIndex        =   67
               Top             =   300
               Width           =   1245
            End
            Begin VB.CommandButton CmdOpciones 
               Caption         =   "Eliminar"
               Enabled         =   0   'False
               Height          =   330
               Index           =   3
               Left            =   2880
               TabIndex        =   10
               Top             =   5940
               Width           =   1305
            End
            Begin VB.CommandButton CmdOpciones 
               Caption         =   "&Modificar"
               Enabled         =   0   'False
               Height          =   330
               Index           =   2
               Left            =   1470
               TabIndex        =   9
               Top             =   5940
               Width           =   1305
            End
            Begin VB.CommandButton CmdOpciones 
               Caption         =   "&Agregar"
               Enabled         =   0   'False
               Height          =   330
               Index           =   1
               Left            =   60
               TabIndex        =   8
               Top             =   5940
               Width           =   1305
            End
            Begin VSFlex7Ctl.VSFlexGrid fg 
               Height          =   5160
               Index           =   3
               Left            =   60
               TabIndex        =   66
               Top             =   660
               Width           =   11595
               _cx             =   20452
               _cy             =   9102
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
               Cols            =   18
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmCronoProduccion3.frx":093E
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
               OutlineCol      =   1
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
               Caption         =   "Semana"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   74
               Top             =   315
               Width           =   585
            End
            Begin VB.Label fchfin 
               AutoSize        =   -1  'True
               Caption         =   "fchfin"
               Height          =   195
               Left            =   4050
               TabIndex        =   72
               Top             =   315
               Width           =   390
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Fin:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   3570
               TabIndex        =   71
               Top             =   315
               Width           =   330
            End
            Begin VB.Label fchini 
               AutoSize        =   -1  'True
               Caption         =   "fchini"
               Height          =   195
               Left            =   2700
               TabIndex        =   70
               Top             =   315
               Width           =   375
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Ini.:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2220
               TabIndex        =   69
               Top             =   315
               Width           =   345
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               Caption         =   "[ Filtro de Fecha ]"
               Height          =   195
               Index           =   2
               Left            =   9000
               TabIndex        =   68
               Top             =   345
               Width           =   1230
            End
            Begin VB.Label LblIdCr 
               AutoSize        =   -1  'True
               Caption         =   "LblIdCr"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   6180
               TabIndex        =   36
               Top             =   6000
               Visible         =   0   'False
               Width           =   495
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle del Programa"
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
            Left            =   60
            TabIndex        =   4
            Top             =   60
            Width           =   11655
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   6690
         Left            =   45
         TabIndex        =   2
         Top             =   390
         Width           =   11760
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6105
            Left            =   30
            TabIndex        =   5
            Top             =   540
            Width           =   11700
            _ExtentX        =   20638
            _ExtentY        =   10769
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
            Columns(3).Caption=   "Programador"
            Columns(3).DataField=   "apenom"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
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
            Splits(0)._ColumnProps(19)=   "Column(3).Width=9102"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=9022"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
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
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
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
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Programa"
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
            TabIndex        =   6
            Top             =   60
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
            Picture         =   "FrmCronoProduccion3.frx":0B2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion3.frx":1073
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion3.frx":11F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion3.frx":164B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion3.frx":1763
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion3.frx":1CA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion3.frx":21EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion3.frx":22FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion3.frx":2413
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion3.frx":2867
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCronoProduccion3.frx":29D3
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
      Width           =   11820
      _ExtentX        =   20849
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
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Linea de Produccion"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
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
   Begin VB.Menu menu2 
      Caption         =   "menu2"
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
      Begin VB.Menu sep0 
         Caption         =   "-"
      End
      Begin VB.Menu menu2_4 
         Caption         =   "Migrar Producto"
      End
   End
   Begin VB.Menu menu3 
      Caption         =   "menu3"
      Visible         =   0   'False
      Begin VB.Menu menu3_1 
         Caption         =   "Activar Seleccionados"
      End
   End
   Begin VB.Menu menu4 
      Caption         =   "menu4"
      Visible         =   0   'False
      Begin VB.Menu menu4_1 
         Caption         =   "Procesar Productos Seleccionados"
         Visible         =   0   'False
      End
      Begin VB.Menu menu4_2 
         Caption         =   "Copiar Seleccionados"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu menu4_3 
         Caption         =   "Migrar Producto"
      End
   End
End
Attribute VB_Name = "FrmCronoProduccion3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
'Dim modifEvent As Boolean
'Dim agregEvent As Boolean
Dim mIdRegistro& 'identificador del registro
Dim OrigFX As Long
Dim OrigFY As Long
Dim DETECTOR_ As CalendarHitTestInfo
Dim EVENTO_ As CalendarEvent
' Variables para las Propiedades de Procesado
Dim MODO_TAREA As Integer ' 0 = "Al finalizar"; 1 = "Al porcentaje"; 2 = "Al minuto"; 3 = "Linea"
Dim PORCENTAJE As Double
Dim MINUTOS_ As String
Dim INCLUIR_HORAS As Integer ' 0 = "Incluir"; 1 = "No incluir"
Dim HOR_INI As String
Dim HOR_FIN As String
Dim LIMITARNUMEROTAREAS_ As Boolean
Dim LIMITARNUMEROPERSONAL_ As Boolean
Dim LIMITARSELPERSONAL_ As Boolean
Dim CORR_ As Double
Dim RstProductos As New ADODB.Recordset
Dim RstProductosAux As New ADODB.Recordset
Dim RstPersonal As New ADODB.Recordset
Dim RstPersonalAux As New ADODB.Recordset
Dim RstTareas As New ADODB.Recordset
Dim RstTareasAux As New ADODB.Recordset

'*****************************************
Dim RstReproc As New ADODB.Recordset
Dim RstReprocAux As New ADODB.Recordset
'*****************************************

Dim cSQL As String
Dim CAMBIO_ As Boolean
Dim ARRASTRANDO_ As Boolean
Dim CARGO_ As Boolean
Dim VERIFICO_ As Boolean
Dim con_SQLS As ADODB.Connection ' Conexion Base de datos del control de asistencia
Dim COLUMNAIDCRDET_ As Integer
Dim COLUMNAIDRECETA_ As Integer
Dim COLUMNAIDITEM_ As Integer
Dim COLUMNAIDLINEA_ As Integer
Dim COLUMNAIDRESP_ As Integer
Dim COLUMNAFCHPROD_ As Integer
Dim COLUMNANUMPROD_ As Integer
Dim COLUMNAPRODUCTO_ As Integer
Dim COLUMNARECETA_ As Integer
Dim COLUMNAUM_ As Integer
Dim COLUMNACANTIDAD_ As Integer
Dim COLUMNAHORINI_ As Integer
Dim COLUMNAHORFIN_ As Integer
Dim COLUMNAFCHFIN_ As Integer
Dim COLUMNAENCARGADO_ As Integer
Dim COLUMNALINEA_ As Integer
Dim COLUMNANUMOPE_ As Integer
Dim COLUMNAPROCESADO_ As Integer
Dim COLUMNACERRADO_ As Integer
Dim COLUMNANUMREGPROD_ As Integer

'******************
Dim COLUMNAPORCENTAJEAPLICADO_ As Integer
Dim COLUMNAREPROC_ As Integer
'******************

Dim ELIMINARTODOS_ As Boolean

'********************************
Dim ESTADO_ As Double

Const ESTADOPENDIENTE_ = 1
Const ESTADOPROCESADO_ = 2
Const ESTADOAPROBADO_ = 3
Const ESTADOANULADO_ = 4
'********************************


'*****************************************************************************************************
'* Descripcion      : EVITA LA EDICION DEL CALENDARIO EN DIVERSAS SITUACIONES
'* Modificacion     : 15/02/11 JOSE CHACON
'*****************************************************************************************************
Private Sub CalCtrlCronog_BeforeEditOperation(ByVal OpParams As XtremeCalendarControl.CalendarEditOperationParameters, bCancelOperation As Boolean)
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
    If OpParams.Operation = xtpCalendarEO_DragMove And EVENTO_.Label <> 9 Then
        ARRASTRANDO_ = True
    Else
        bCancelOperation = True
        ARRASTRANDO_ = False
    End If
    ' Eliminacion Manual
    If OpParams.Operation = xtpCalendarEO_DeleteEvent Then bCancelOperation = True
End Sub

Private Sub CalCtrlCronog_DblClick()
    Agregando = True
    mostrarFormulario False, True, False
    Agregando = False
    Set DETECTOR_ = Nothing
End Sub

Private Sub CalCtrlCronog_KeyDown(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    
On Error Resume Next
    ' Se activa el detector para la vista activa del calendario
    Set DETECTOR_ = CalCtrlCronog.ActiveView.HitTest
    
    ' Se agrega el evento del detector
    Set EVENTO_ = DETECTOR_.ViewEvent.Event
    
    If KeyCode = vbKeyInsert Then
        Menu2_1_Click
    End If
    
    If KeyCode = vbKeyDelete Then
        ' Si el detector no tiene evento activo
        If DETECTOR_.ViewEvent Is Nothing Then Exit Sub
        menu2_2_Click
    End If
    
    If KeyCode = 113 Then
        ' Si el detector no tiene evento activo
        If DETECTOR_.ViewEvent Is Nothing Then Exit Sub
        Menu2_3_Click
    End If
    
    Set DETECTOR_ = Nothing
End Sub

Private Sub CalCtrlCronog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim xRs As New ADODB.Recordset
    Dim IDEVENTO_ As Double
    Dim Rpta As Integer
    Dim EVENTO_AUX As CalendarEvent
    Dim EVENTOINICIAL_ As CalendarEvent
    Dim CANTIDAD_ As Double
    Dim IDLINEA_ As Double
    Dim IDCRDET_ As Double
    Dim IDITEM_ As Double
    Dim HORINI_ As String
    Dim FECHINI_ As Date
    
On Error GoTo ERROR_
            
    If ARRASTRANDO_ Then
        IDEVENTO_ = EVENTO_.ID
        Set EVENTOINICIAL_ = EVENTO_
         
        Set EVENTO_ = CalCtrlCronog.DataProvider.GetEvent(IDEVENTO_)
                
        If EVENTOINICIAL_ <> EVENTO_ Then
            RstProductos.Filter = adFilterNone
            ' Se limpia el calendario
            CalCtrlCronog.DataProvider.RemoveAllEvents
            ' Se llenan todos los eventos sin modificar
            llenarDatos
            Exit Sub
        End If
    
        ' Se llena el evento auxiliar
        Set EVENTO_AUX = EVENTO_
        ' Se elimina el evento arrastrado
        CalCtrlCronog.DataProvider.DeleteEvent EVENTO_
        
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
        IDLINEA_ = NulosN(RstProductosAux("idlinea"))
        IDCRDET_ = NulosN(RstProductosAux("id"))
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
            CalCtrlCronog.DataProvider.RemoveAllEvents
            ' Se llenan todos los eventos
            llenarDatos False, IDCRDET_
        Else
            RstProductos.Filter = adFilterNone
            ' Se limpia el calendario
            CalCtrlCronog.DataProvider.RemoveAllEvents
            ' Se llenan todos los eventos sin modificar
            llenarDatos False, IDCRDET_
        End If
    End If
    
    ARRASTRANDO_ = False
    Exit Sub
ERROR_:
    MsgBox "Ocurrio un error al desplazar el evento; intente de nuevo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Sub

Private Sub CalCtrlCronog_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Set DETECTOR_ = CalCtrlCronog.ActiveView.HitTest
    Set EVENTO_ = DETECTOR_.ViewEvent.Event
        
    If Button = 2 Then
        If QueHace <> 3 Then PopupMenu menu2
    End If
End Sub

Private Sub CalCtrlCronog_SelectionChanged(ByVal SelType As XtremeCalendarControl.CalendarSelectionChanged)
    If SelType = xtpCalendarSelectionDays Then
        If Agregando Then Exit Sub
        If Not CARGO_ Then Exit Sub ' Si no ha cargado el calendario
        If VERIFICO_ Then Exit Sub ' Si se verfico que corresponde al rango de fechas

        Dim FCHINI_ As Date
        Dim FCHFIN_ As Date
        Dim TODODIA_ As Boolean
        Dim PRIMERDIASEMANA_ As Date
        Dim ULTIMODIASEMANA_ As Date

        ' Se obtienen los datos del dia seleccionado
        CalCtrlCronog.ActiveView.GetSelection FCHINI_, FCHFIN_, TODODIA_

        ' SI es una fecha Incoherente
        If Format(FCHINI_, "yyyy") < AnoTra Then Exit Sub

        PRIMERDIASEMANA_ = CDate(TxtFchIni.valor)
        ULTIMODIASEMANA_ = CDate(TxtFchFin.valor)
        FCHINI_ = Format(FCHINI_, "dd/mm/yyyy")

        If FCHINI_ < PRIMERDIASEMANA_ Or FCHINI_ > ULTIMODIASEMANA_ Then
            CalCtrlCronog.ActiveView.ShowDay (PRIMERDIASEMANA_)
            VERIFICO_ = True
            CalCtrlCronog.ViewType = xtpCalendarFullWeekView
        End If
        VERIFICO_ = False
    End If
End Sub

Sub CrearCabeceraVS(numPag As Integer)
    Dim xCad As String

    FrmVsPrinter.Vs.TextAlign = taLeftTop
    FrmVsPrinter.Vs.FontName = "Courier New"
    FrmVsPrinter.Vs.FontBold = True
    FrmVsPrinter.Vs.FontSize = 9

    FrmVsPrinter.Vs.CurrentX = 1000:      FrmVsPrinter.Vs.CurrentY = 200
    FrmVsPrinter.Vs.Paragraph = "EMPRESA   : " & NomEmp

    FrmVsPrinter.Vs.CurrentX = 7950:      FrmVsPrinter.Vs.CurrentY = 200
    FrmVsPrinter.Vs.Paragraph = "FECHA        : " & Format(Date, "dd/mm/yy")

    FrmVsPrinter.Vs.CurrentX = 1000:      FrmVsPrinter.Vs.CurrentY = 400
    FrmVsPrinter.Vs.Paragraph = "Nº R.U.C. : " & NumRUC

    FrmVsPrinter.Vs.CurrentX = 7950:      FrmVsPrinter.Vs.CurrentY = 400
    FrmVsPrinter.Vs.Paragraph = "Nº Pagina    : " & Format(numPag, "0000")

    FrmVsPrinter.Vs.DrawLine 1000, 650, 11000, 650
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

Private Sub imprimirReporte()
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios
    Dim TITULO_ As String
    
    TITULO_ = "REPORTE DE CRONOGRAMA DE PRODUCCION " & TxtFchIni.valor & " - " & TxtFchFin.valor

    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, fg(3), TITULO_, "", "", TITULO_
    Set X_EXPORT = Nothing
    MousePointer = vbDefault
    Exit Sub
error:
    MousePointer = vbDefault
    SHOW_ERROR Name, "Imprimir Reporte"
End Sub

Private Sub imprimirDetallado(IDCRDET_ As Double, IDAREA_ As Double, IDRESP_ As Double, ByRef numPag As Integer, ByRef RECORDSET_ As ADODB.Recordset)
    Dim RstTarAux As New ADODB.Recordset
    Dim HORINI_ As String
    Dim HORFIN_ As String
    Dim CANT_ As Double
    Dim CAMBIO_ As Boolean
    Dim A As Integer
    Dim xRsTarAuxAux As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim xLinea As Double
    Dim B As Integer
    Dim xColumna As Integer         ' Columna de impresion
    Dim numper As Double            ' Numero de Personas
    Dim ID_LINEA As Double
    Dim RESPONSABLE_ As String
    
    DEFINIR_RST_TMP RstTarAux, RECORDSET_
    CARGAR_RST_TMP RstTarAux, RECORDSET_
        
    RstTarAux.Filter = "idarea=" & IDAREA_ & " And idresp=" & IDRESP_ & " And activo=-1"
    
    ' Se graba las tareas
    RstTarAux.MoveFirst
    RESPONSABLE_ = NulosC(RstTarAux("nomresp"))
    HORINI_ = Format(RstTarAux("horinitar"), "HH:mm")
    
    RstTarAux.MoveLast
    HORFIN_ = Format(RstTarAux("horfintar"), "HH:mm")
    CANT_ = NulosN(RstTarAux("cantproc"))
                        
    Set xRs = definirUnirPersonal(RstTarAux, RstPersonal)
    
    With FrmVsPrinter.Vs
        xLinea = 1300
        xColumna = 900
        numPag = numPag + 1
        CrearCabeceraVS numPag
        
        RstProductos.Filter = "id=" & IDCRDET_
        If RstProductos.RecordCount = 0 Then Exit Sub
        '******************************************************************* Titulo
        .FontSize = 12
        .FontBold = True
        .TextAlign = taCenterMiddle
        
        .TextBox "LINEA DE PRODUCCION", xColumna, xLinea, 8000, 500, True, False, True
        .FontSize = 10
        .TextAlign = taCenterTop
        .TextBox "Nº REG. PROG.", xColumna + 8100, xLinea, 1900, 250, True, False, True
        xLinea = xLinea + 240
        .TextBox NulosC(RstProductos("numregprod")), xColumna + 8100, xLinea, 1900, 250, True, False, True
        
        .TextAlign = taLeftMiddle
        .FontSize = 9
        
        .FontBold = False
        xLinea = xLinea + 250
        .TextBox "Producto", xColumna, xLinea, 1500, 250, True, False, False
        .TextBox NulosC(RstProductos("descripcion")), xColumna + 1500, xLinea, 7000, 250, True, False, False
        
        .TextBox "Receta", xColumna + 7500, xLinea, 1000, 250, True, False, False
        .TextBox NulosC(RstProductos("codrec")), xColumna + 8550, xLinea, 6000, 250, True, False, False
        
        '*************************************************************************
        xLinea = xLinea + 250
        .TextBox "Fecha Prog.", xColumna, xLinea, 1500, 250, True, False, False
        .TextBox Format(RstProductos("fchpro"), "dd/mm/yyyy"), xColumna + 1500, xLinea, 6000, 250, True, False, False
                    
        '*************************************************************************
        .TextBox "Cantidad", xColumna + 7500, xLinea, 1000, 250, True, False, False
        .TextBox Format(CANT_, "0.00") & " " & encontrarUnidad(RstProductos("iditem")), xColumna + 8550, xLinea, 6000, 250, True, False, False
        
        '*************************************************************************
        xLinea = xLinea + 250
        .TextBox "Responsable ", xColumna, xLinea, 1500, 250, True, False, False
        .TextBox RESPONSABLE_, xColumna + 1500, xLinea, 6000, 250, True, False, False
            
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
        
        RstTarAux.MoveFirst
        
'        Dim xRsTarAux As New ADODB.Recordset
'
'        DEFINIR_RST_TMP xRsTarAux, RstTarAux
'        CARGAR_RST_TMP xRsTarAux, RstTarAux
            
        ID_LINEA = NulosN(RstProductos("idlinea"))
        
'        xRsTarAux.MoveFirst
        For B = 1 To RstTarAux.RecordCount
            .FontSize = 8
            .FontBold = False
            
            .TextAlign = taLeftMiddle
            .TextBox " " & NulosN(RstTarAux("idtar")), xColumna, xLinea, 500, 250, True, False, True
            .TextBox " " & NulosC(RstTarAux("destar")), xColumna + 500, xLinea, 3500, 250, True, False, True
            .TextAlign = taCenterMiddle
            .TextBox Format(RstTarAux("durtar"), "HH:mm"), xColumna + 4000, xLinea, 800, 250, True, False, True
            .TextBox Format(RstTarAux("horinitar"), "HH:mm"), xColumna + 4800, xLinea, 800, 250, True, False, True
            .TextBox Format(RstTarAux("horfintar"), "HH:mm"), xColumna + 5600, xLinea, 800, 250, True, False, True
            .TextBox Format(RstTarAux("numper"), "00"), xColumna + 6400, xLinea, 800, 250, True, False, True
            
            .TextAlign = taRightMiddle
            .TextBox Format(hallarCaracTareas(ID_LINEA, RstTarAux("idtar")), "0.00") & " ", xColumna + 7200, xLinea, 1000, 250, True, False, True
            .TextBox Format(RstTarAux("aplpor"), "0.00") & "% ", xColumna + 8200, xLinea, 800, 250, True, False, True
            .TextBox Format(RstTarAux("cantproc"), "0.00") & " ", xColumna + 9000, xLinea, 1000, 250, True, False, True
            
            numper = numper + NulosN(RstTarAux("numper"))
            
            RstTarAux.MoveNext
            If RstTarAux.EOF = True Then Exit For
            
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
        .TextBox "RECETA", xColumna + 4750, xLinea, 1500, 250, True, False, True
        .TextBox "CANTIDAD", xColumna + 6250, xLinea, 1000, 250, True, False, True
        
        .TextAlign = taCenterMiddle
        xLinea = xLinea + 250
        .FontSize = 7
        .TextBox calcularProdAnterior(ID_LINEA, True, True), xColumna + 500, xLinea, 4250, 250, True, False, True
        .FontSize = 8
        .TextBox "", xColumna + 4750, xLinea, 1500, 250, True, False, True
        .TextBox "", xColumna + 6250, xLinea, 1000, 250, True, False, True
        
        .TextAlign = taLeftMiddle
        .TextBox " Hora Ini.", xColumna + 7500, xLinea, 1500, 250, True, False, True
        .TextBox "", xColumna + 8500, xLinea, 1500, 250, True, False, True
        '*************************************************************************
        
        .TextAlign = taCenterMiddle
        xLinea = xLinea + 250
        .TextBox "P1", xColumna, xLinea, 500, 250, True, False, True
        .FontSize = 7
        .TextBox NulosC(RstProductos("descripcion")), xColumna + 500, xLinea, 4250, 250, True, False, True
        .FontSize = 8
        .TextBox NulosC(RstProductos("codrec")), xColumna + 4750, xLinea, 1500, 250, True, False, True
        .TextBox "", xColumna + 6250, xLinea, 1000, 250, True, False, True
        
        .TextAlign = taLeftMiddle
        .TextBox " Hora Fin", xColumna + 7500, xLinea, 1500, 250, True, False, True
        .TextBox "", xColumna + 8500, xLinea, 1500, 250, True, False, True
        '*************************************************************************
        xLinea = xLinea + 250
        .TextAlign = taCenterMiddle
        .TextBox "P2", xColumna, xLinea, 500, 250, True, False, True
        .TextBox "", xColumna + 500, xLinea, 4250, 250, True, False, True
        .TextBox "", xColumna + 4750, xLinea, 1500, 250, True, False, True
        .TextBox "", xColumna + 6250, xLinea, 1000, 250, True, False, True
        '*************************************************************************
        xLinea = xLinea + 250
        .TextBox "P3", xColumna, xLinea, 500, 250, True, False, True
        .TextBox "", xColumna + 500, xLinea, 4250, 250, True, False, True
        .TextBox "", xColumna + 4750, xLinea, 1500, 250, True, False, True
        .TextBox "", xColumna + 6250, xLinea, 1000, 250, True, False, True
        
        .TextAlign = taLeftMiddle
        .TextBox " Lote", xColumna + 7500, xLinea, 1500, 250, True, False, True
        .TextBox "", xColumna + 8500, xLinea, 1500, 250, True, False, True
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
        .TextBox "Prod1", xColumna + 7600, xLinea, 600, 500, True, False, True
        .TextBox "Prod2", xColumna + 8200, xLinea, 600, 500, True, False, True
        .TextBox "Prod3", xColumna + 8800, xLinea, 600, 500, True, False, True
        .TextBox "Efic.", xColumna + 9400, xLinea, 600, 500, True, False, True
            
        'numper = xRs.RecordCount
            
        If xRs.RecordCount > 0 Then xRs.MoveFirst
        
        xLinea = xLinea + 500
        xFila = xLinea
        
        ' Se agrega 5 campos mas para ingresar personal
        numper = numper + 5
        ' Se verifica que se muestre no menos de 25 personas
        For B = 1 To numper
            .FontSize = 10
            .FontBold = False
            .TextAlign = taLeftMiddle
            
            .TextBox " " & Format(B, "00"), xColumna, xLinea, 500, 300, True, False, True
            If Not xRs.EOF Then
                .TextBox " " & NulosC(xRs("nombre")), xColumna + 500, xLinea, 3500, 300, True, False, True
                .TextBox " " & NulosC(xRs("idtar")), xColumna + 4000, xLinea, 800, 300, True, False, True
                .TextBox "", xColumna + 4800, xLinea, 1000, 300, True, False, True
                .TextBox "", xColumna + 5800, xLinea, 1000, 300, True, False, True
                .TextBox "", xColumna + 6800, xLinea, 800, 300, True, False, True
                .TextBox "", xColumna + 7600, xLinea, 600, 300, True, False, True
                .TextBox "", xColumna + 8200, xLinea, 600, 300, True, False, True
                .TextBox "", xColumna + 8800, xLinea, 600, 300, True, False, True
                .TextBox "", xColumna + 9400, xLinea, 600, 300, True, False, True
                xRs.MoveNext
            Else
                .TextBox "", xColumna + 500, xLinea, 3500, 300, True, False, True
                .TextBox "", xColumna + 4000, xLinea, 800, 300, True, False, True
                .TextBox "", xColumna + 4800, xLinea, 1000, 300, True, False, True
                .TextBox "", xColumna + 5800, xLinea, 1000, 300, True, False, True
                .TextBox "", xColumna + 6800, xLinea, 800, 300, True, False, True
                .TextBox "", xColumna + 7600, xLinea, 600, 300, True, False, True
                .TextBox "", xColumna + 8200, xLinea, 600, 300, True, False, True
                .TextBox "", xColumna + 8800, xLinea, 600, 300, True, False, True
                .TextBox "", xColumna + 9400, xLinea, 600, 300, True, False, True
            End If
            
            xLinea = xLinea + 300
            
            If xLinea >= 14750 Then
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
        
    End With
End Sub

Private Sub ImprimirLinea(IDCRDET_ As Double)
    Dim A As Integer
    Dim B As Integer
    Dim xLinea As Integer           ' Fila de impresion
    Dim xColumna As Integer         ' Columna de impresion
    Dim numPag As Integer           ' Numero de pagina de Impresion
    Dim numper As Double            ' Numero de Personas
    Dim ID_LINEA As Double
    Dim Rst As New ADODB.Recordset
    Dim IDAREA_ As Double
    Dim IDRESP_ As Double
    
    With FrmVsPrinter.Vs
        .BrushColor = &H80000005
        .FontSize = 11
        .TextAlign = taCenterMiddle
            
        'On Error Resume Next
                
        If RstTareas.State = 0 Then Exit Sub
        RstTareas.Filter = "idcrdet=" & IDCRDET_ & " And activo=-1"
        If RstTareas.RecordCount = 0 Then Exit Sub
        
        DEFINIR_RST_TMP Rst, RstTareas
        CARGAR_RST_TMP Rst, RstTareas
        
        Rst.MoveFirst
        IDAREA_ = NulosN(Rst("idarea"))
        IDRESP_ = NulosN(Rst("idresp"))
        numPag = 0
        CAMBIO_ = True
        
        RstTareas.MoveFirst
        While Not RstTareas.EOF
            If Not CAMBIO_ Then GoTo SIGUIENTE_
            imprimirDetallado IDCRDET_, IDAREA_, IDRESP_, numPag, Rst
            
SIGUIENTE_:
            RstTareas.MoveNext
            If Not RstTareas.EOF Then
                If IDAREA_ <> NulosN(RstTareas("idarea")) Or IDRESP_ <> NulosN(RstTareas("idresp")) Then
                    CAMBIO_ = True
                    numPag = numPag + 1
                    IDAREA_ = NulosN(RstTareas("idarea"))
                    IDRESP_ = NulosN(RstTareas("idresp"))
                    .NewPage
                    CrearCabeceraVS numPag
                Else
                    CAMBIO_ = False
                End If
            End If
        Wend
        
'
'        xLinea = 1300
'        xColumna = 900
'        numPag = numPag + 1
'        If A > 1 Then .NewPage
'        CrearCabeceraVS numPag
'
'        '******************************************************************* Titulo
'        .FontSize = 12
'        .FontBold = True
'        .TextAlign = taCenterMiddle
'
'        .TextBox "LINEA DE PRODUCCION", xColumna, xLinea, 8000, 500, True, False, True
'        .FontSize = 10
'        .TextAlign = taCenterTop
'        .TextBox "NUM. PROG.", xColumna + 8100, xLinea, 1900, 250, True, False, True
'        xLinea = xLinea + 240
'        .TextBox NulosC(RstProductos("numprod")), xColumna + 8100, xLinea, 1900, 250, True, False, True
'
'        .TextAlign = taLeftMiddle
'        .FontSize = 9
'
'        .FontBold = False
'        xLinea = xLinea + 250
'        .TextBox "Producto", xColumna, xLinea, 1500, 250, True, False, False
'        .TextBox NulosC(RstProductos("descripcion")), xColumna + 1500, xLinea, 7000, 250, True, False, False
'
'        .TextBox "Receta", xColumna + 7500, xLinea, 1000, 250, True, False, False
'        .TextBox NulosC(RstProductos("codrec")), xColumna + 8550, xLinea, 6000, 250, True, False, False
'
'        '*************************************************************************
'        xLinea = xLinea + 250
'        .TextBox "Fecha Prog.", xColumna, xLinea, 1500, 250, True, False, False
'        .TextBox Format(RstProductos("fchpro"), "dd/mm/yyyy"), xColumna + 1500, xLinea, 6000, 250, True, False, False
'
'        '*************************************************************************
'        .TextBox "Cantidad", xColumna + 7500, xLinea, 1000, 250, True, False, False
'        .TextBox Format(RstProductos("cantidad"), "0.00") & " " & encontrarUnidad(RstProductos("iditem")), xColumna + 8550, xLinea, 6000, 250, True, False, False
'
'        '*************************************************************************
'        xLinea = xLinea + 250
'        .TextBox "Responsable ", xColumna, xLinea, 1500, 250, True, False, False
'        .TextBox RstProductos("nomresp"), xColumna + 1500, xLinea, 6000, 250, True, False, False
'
'        Dim xFila As Integer
'        '******************************************************************* Detalle de la Linea
'        xLinea = xLinea + 300
'        .TextAlign = taLeftMiddle
'        .FontBold = True
'        .TextBox "Detalles de la Linea", xColumna, xLinea, 2500, 250, True, False, False
'        '*************************************************************************
'
'        .FontBold = False
'        xLinea = xLinea + 350
'        .TextAlign = taCenterMiddle
'        .TextBox "Ord.", xColumna, xLinea, 500, 500, True, False, True
'        .TextBox "Tarea", xColumna + 500, xLinea, 3500, 500, True, False, True
'        .TextBox "Durac.", xColumna + 4000, xLinea, 800, 500, True, False, True
'        .TextBox "Hor.Ini", xColumna + 4800, xLinea, 800, 500, True, False, True
'        .TextBox "Hor.Fin", xColumna + 5600, xLinea, 800, 500, True, False, True
'        .TextBox "Num. Pers.", xColumna + 6400, xLinea, 800, 500, True, False, True
'        .TextBox "Unid.x Hora", xColumna + 7200, xLinea, 1000, 500, True, False, True
'        .TextBox "%Rdto", xColumna + 8200, xLinea, 800, 500, True, False, True
'        .TextBox "Cant. Proc.", xColumna + 9000, xLinea, 1000, 500, True, False, True
'
'        numper = 0
'        xLinea = xLinea + 500
'        xFila = xLinea
'
'        RstTareas.Filter = "idcrdet = " & IDCRDET_ & " And activo = True"
'        RstTareas.MoveFirst
'
'        Dim xRsTarAux As New ADODB.Recordset
'
'        DEFINIR_RST_TMP xRsTarAux, RstTareas
'        CARGAR_RST_TMP xRsTarAux, RstTareas
'
'
'        ID_LINEA = NulosN(RstProductos("idlinea"))
'
'        For B = 1 To RstTareas.RecordCount
'            .FontSize = 8
'            .FontBold = False
'
'            .TextAlign = taLeftMiddle
'            .TextBox " " & NulosN(RstTareas("idtar")), xColumna, xLinea, 500, 250, True, False, True
'            .TextBox " " & NulosC(RstTareas("destar")), xColumna + 500, xLinea, 3500, 250, True, False, True
'            .TextAlign = taCenterMiddle
'            .TextBox Format(RstTareas("durtar"), "HH:mm"), xColumna + 4000, xLinea, 800, 250, True, False, True
'            .TextBox Format(RstTareas("horinitar"), "HH:mm"), xColumna + 4800, xLinea, 800, 250, True, False, True
'            .TextBox Format(RstTareas("horfintar"), "HH:mm"), xColumna + 5600, xLinea, 800, 250, True, False, True
'            .TextBox Format(RstTareas("numper"), "00"), xColumna + 6400, xLinea, 800, 250, True, False, True
'
'            .TextAlign = taRightMiddle
'            .TextBox Format(hallarCaracTareas(ID_LINEA, RstTareas("idtar")), "0.00") & " ", xColumna + 7200, xLinea, 1000, 250, True, False, True
'            .TextBox Format(RstTareas("aplpor"), "0.00") & "% ", xColumna + 8200, xLinea, 800, 250, True, False, True
'            .TextBox Format(RstTareas("cantproc"), "0.00") & " ", xColumna + 9000, xLinea, 1000, 250, True, False, True
'
'            numper = numper + NulosN(RstTareas("numper"))
'
'            RstTareas.MoveNext
'            If RstTareas.EOF = True Then Exit For
'
'            xLinea = xLinea + 250
'
'            If xLinea >= 16200 Then
'                xLinea = 1300
'                numPag = numPag + 1
'                .NewPage
'                CrearCabeceraVS numPag
'            End If
'        Next B
'
'        xLinea = xLinea + 250
'        .TextAlign = taRightMiddle
'        .TextBox "TOTAL", xColumna, xLinea, 4000, 250, True, False, True
'        .TextAlign = taCenterMiddle
'        .TextBox Format(numper, "00"), xColumna + 6400, xLinea, 800, 250, True, False, True
'
'        .FontBold = False
'        xLinea = xLinea + 400
'        .TextBox "RECETA", xColumna + 4750, xLinea, 1500, 250, True, False, True
'        .TextBox "CANTIDAD", xColumna + 6250, xLinea, 1000, 250, True, False, True
'
'        .TextAlign = taCenterMiddle
'        xLinea = xLinea + 250
'        .FontSize = 7
'        .TextBox calcularProdAnterior(ID_LINEA, True, True), xColumna + 500, xLinea, 4250, 250, True, False, True
'        .FontSize = 8
'        .TextBox "", xColumna + 4750, xLinea, 1500, 250, True, False, True
'        .TextBox "", xColumna + 6250, xLinea, 1000, 250, True, False, True
'
'        .TextAlign = taLeftMiddle
'        .TextBox " Hora Ini.", xColumna + 7500, xLinea, 1500, 250, True, False, True
'        .TextBox "", xColumna + 8500, xLinea, 1500, 250, True, False, True
'        '*************************************************************************
'
'        .TextAlign = taCenterMiddle
'        xLinea = xLinea + 250
'        .TextBox "P1", xColumna, xLinea, 500, 250, True, False, True
'        .FontSize = 7
'        .TextBox NulosC(RstProductos("descripcion")), xColumna + 500, xLinea, 4250, 250, True, False, True
'        .FontSize = 8
'        .TextBox NulosC(RstProductos("codrec")), xColumna + 4750, xLinea, 1500, 250, True, False, True
'        .TextBox "", xColumna + 6250, xLinea, 1000, 250, True, False, True
'
'        .TextAlign = taLeftMiddle
'        .TextBox " Hora Fin", xColumna + 7500, xLinea, 1500, 250, True, False, True
'        .TextBox "", xColumna + 8500, xLinea, 1500, 250, True, False, True
'        '*************************************************************************
'        xLinea = xLinea + 250
'        .TextAlign = taCenterMiddle
'        .TextBox "P2", xColumna, xLinea, 500, 250, True, False, True
'        .TextBox "", xColumna + 500, xLinea, 4250, 250, True, False, True
'        .TextBox "", xColumna + 4750, xLinea, 1500, 250, True, False, True
'        .TextBox "", xColumna + 6250, xLinea, 1000, 250, True, False, True
'        '*************************************************************************
'        xLinea = xLinea + 250
'        .TextBox "P3", xColumna, xLinea, 500, 250, True, False, True
'        .TextBox "", xColumna + 500, xLinea, 4250, 250, True, False, True
'        .TextBox "", xColumna + 4750, xLinea, 1500, 250, True, False, True
'        .TextBox "", xColumna + 6250, xLinea, 1000, 250, True, False, True
'
'        .TextAlign = taLeftMiddle
'        .TextBox " Lote", xColumna + 7500, xLinea, 1500, 250, True, False, True
'        .TextBox "", xColumna + 8500, xLinea, 1500, 250, True, False, True
'        '*************************************************************************
'
'        '****************************************************************************************
'        '******************************************************************* Detalle del Personal
'        '****************************************************************************************
'        xLinea = xLinea + 300
'        .TextAlign = taLeftMiddle
'        .FontBold = True
'        .TextBox "Detalles del Personal", xColumna, xLinea, 2500, 250, True, False, False
'        '*************************************************************************
'
'        xLinea = xLinea + 350
'
'        .FontBold = False
'        .TextAlign = taCenterMiddle
'        .TextBox "Item", xColumna, xLinea, 500, 500, True, False, True
'        .TextBox "PERSONAL", xColumna + 500, xLinea, 3500, 500, True, False, True
'        .TextBox "Tarea", xColumna + 4000, xLinea, 800, 500, True, False, True
'        .TextBox "Hr.Ini.", xColumna + 4800, xLinea, 1000, 500, True, False, True
'        .TextBox "Hr.Ter.", xColumna + 5800, xLinea, 1000, 500, True, False, True
'        .TextBox "M.P.", xColumna + 6800, xLinea, 800, 500, True, False, True
'        .TextBox "Prod1", xColumna + 7600, xLinea, 600, 500, True, False, True
'        .TextBox "Prod2", xColumna + 8200, xLinea, 600, 500, True, False, True
'        .TextBox "Prod3", xColumna + 8800, xLinea, 600, 500, True, False, True
'        .TextBox "Efic.", xColumna + 9400, xLinea, 600, 500, True, False, True
'
'        RstPersonal.Filter = "idcrdet = " & IDCRDET_
'        If RstPersonal.RecordCount <> 0 Then RstPersonal.MoveFirst
'
'        xLinea = xLinea + 500
'        xFila = xLinea
'
'        ' Se agrega 5 campos mas para ingresar personal
'        numper = numper + 5
'        ' Se verifica que se muestre no menos de 25 personas
'        For B = 1 To numper
'            .FontSize = 10
'            .FontBold = False
'            .TextAlign = taLeftMiddle
'
'            .TextBox " " & Format(B, "00"), xColumna, xLinea, 500, 300, True, False, True
'            If Not RstPersonal.EOF Then
'                .TextBox " " & NulosC(RstPersonal("nombre")), xColumna + 500, xLinea, 3500, 300, True, False, True
'                .TextBox " " & NulosC(RstPersonal("idtar")), xColumna + 4000, xLinea, 800, 300, True, False, True
'                .TextBox "", xColumna + 4800, xLinea, 1000, 300, True, False, True
'                .TextBox "", xColumna + 5800, xLinea, 1000, 300, True, False, True
'                .TextBox "", xColumna + 6800, xLinea, 800, 300, True, False, True
'                .TextBox "", xColumna + 7600, xLinea, 600, 300, True, False, True
'                .TextBox "", xColumna + 8200, xLinea, 600, 300, True, False, True
'                .TextBox "", xColumna + 8800, xLinea, 600, 300, True, False, True
'                .TextBox "", xColumna + 9400, xLinea, 600, 300, True, False, True
'                RstPersonal.MoveNext
'            Else
'                .TextBox "", xColumna + 500, xLinea, 3500, 300, True, False, True
'                .TextBox "", xColumna + 4000, xLinea, 800, 300, True, False, True
'                .TextBox "", xColumna + 4800, xLinea, 1000, 300, True, False, True
'                .TextBox "", xColumna + 5800, xLinea, 1000, 300, True, False, True
'                .TextBox "", xColumna + 6800, xLinea, 800, 300, True, False, True
'                .TextBox "", xColumna + 7600, xLinea, 600, 300, True, False, True
'                .TextBox "", xColumna + 8200, xLinea, 600, 300, True, False, True
'                .TextBox "", xColumna + 8800, xLinea, 600, 300, True, False, True
'                .TextBox "", xColumna + 9400, xLinea, 600, 300, True, False, True
'            End If
'
'            xLinea = xLinea + 300
'
'            If xLinea >= 14750 Then
'                xLinea = 1300
'                numPag = numPag + 1
'                .NewPage
'                CrearCabeceraVS numPag
'            End If
'        Next B
'
'        '****************************************************************************************
'        '************************************************************************** Observaciones
'        '****************************************************************************************
'        xLinea = xLinea + 100
'
'        If xLinea >= 15500 Then
'            xLinea = 1300
'            numPag = numPag + 1
'            .NewPage
'            CrearCabeceraVS numPag
'        End If
'
'        .TextAlign = taLeftMiddle
'        .FontBold = True
'        .TextBox "Observaciones", xColumna, xLinea, 2500, 250, True, False, False
'        '*************************************************************************
'        xLinea = xLinea + 450
'        .DrawLine xColumna + 500, xLinea, 10000, xLinea
'        xLinea = xLinea + 250
'        .DrawLine xColumna + 500, xLinea, 10000, xLinea
'        xLinea = xLinea + 250
'        .DrawLine xColumna + 500, xLinea, 10000, xLinea
'        xLinea = xLinea + 250
'        .DrawLine xColumna + 500, xLinea, 10000, xLinea
SIGUIENTE:
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
            cSQL = "SELECT 0 AS xsel, pro_cronogramadet.fchpro, alm_inventario.descripcion, pro_cronogramadet.cantidad, pro_cronogramadet.horpro, pro_cronogramadet.horfin, pro_cronogramadet.id, pro_cronogramadet.idcr, pro_cronograma.semana, pla_empleados.nombre, pro_cronogramadet.estado " _
                + vbCr + "FROM (pro_cronograma RIGHT JOIN (pro_cronogramadet LEFT JOIN alm_inventario ON pro_cronogramadet.iditem = alm_inventario.id) ON pro_cronograma.id = pro_cronogramadet.idcr) LEFT JOIN pla_empleados ON pro_cronogramadet.idresp = pla_empleados.id " _
                + vbCr + "WHERE (((pro_cronograma.semana) = " & NulosN(ComboSemanas.Text) & ") And ((pro_cronogramadet.estado) = " & ESTADOPROCESADO_ & ")) " _
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
                    ImprimirLinea NulosN(xRs("id"))
SIGUIENTE:
                    xRs.MoveNext
                Next A
                Me.MousePointer = vbDefault
                .EndDoc
            End With
            'Muestra la preimagen de la impresion
            FrmVsPrinter.WindowState = 2
            FrmVsPrinter.Show
            
        Case 1
            If QueHace <> 3 Then Exit Sub
            If CalCtrlCronog.Visible Then CmdOpciones_Click 5
            Agregando = True
            imprimirReporte
            Agregando = False
            
    End Select
End Sub

Private Sub iniciarCampos()
    Dim A As Integer
    'Dim pTema2007 As CalendarThemeOffice2007
    
    'Se guarda el tema del calendario activo
    'Set pTema2007 = CalCtrlCronog.Theme

    'Se cambia el color de seleccion
'    pTema2007.Event.Normal.Location.Color = &HFF&
'
'    pTema2007.Event.Normal.Body.Color = &HFF0000
'    pTema2007.Event.Normal.Body.Font.Size = 8
'    pTema2007.Event.Selected.Background.ColorDark = &HFF0000
'    pTema2007.Event.Selected.BorderColor = &HFF&
'
'    pTema2007.Event.Selected.Subject.Color = &HFF&
'    pTema2007.Event.Selected.Subject.Font.Size = 7
'    pTema2007.Event.Normal.Subject.Font.Size = 7
'    pTema2007.Event.Normal.Subject.Font.Bold = True

    ' Se habilita los mensajes de ayuda
'    CalCtrlCronog.EnableToolTips True
'    ' Se deshabilita el ingreso de eventos por mouse
'    CalCtrlCronog.Options.EnableAddNewTooltip = False
    
    SliderCal.Max = 1000
    SliderCal.Min = 150
    SliderCal.TickFrequency = 100
    SliderCal.Value = 400
    
'    CalCtrlCronog.DayView.TimeScale = 20
'    CalCtrlCronog.DayView.EnableHScroll False
'    CalCtrlCronog.DayView.MinColumnWidth = SliderCal.Value
'    CalCtrlCronog.Options.DayViewTimeScaleShowMinutes = True
    

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
    ' ****************************************
    LIMITARNUMEROTAREAS_ = True
    LIMITARNUMEROPERSONAL_ = True
    LIMITARSELPERSONAL_ = True
    ' ****************************************
    CORR_ = -666
    
    fg(0).AllowUserResizing = flexResizeColumns
    'fg(0).ColWidth(8) = 0
    fg(0).ColWidth(9) = 0
    fg(0).ColWidth(10) = 0
    fg(0).ColWidth(11) = 0
    fg(0).ColWidth(12) = 0
    fg(0).ColWidth(14) = 0
    fg(0).FrozenCols = 2
    fg(0).ColWidth(16) = 0
    fg(0).ColWidth(17) = 0
    fg(0).ColWidth(18) = 0
    GRID_COMBOLIST fg(0), 13
    GRID_COMBOLIST fg(0), 14
    GRID_COMBOLIST fg(0), 15
    
    'fg(1).ColWidth(1) = 0
'    fg(1).ColWidth(4) = 0
'    fg(1).ColWidth(5) = 0
'    fg(1).ColWidth(6) = 0
'    fg(1).ColWidth(7) = 0
    
    fg(2).ColWidth(0) = 0
    fg(2).ColWidth(5) = 0
    fg(2).ColWidth(6) = 0
    fg(2).ColWidth(7) = 0
        
    fg(3).AllowUserResizing = flexResizeColumns
    fg(3).ExplorerBar = flexExSortShow
    fg(3).SelectionMode = flexSelectionFree
    fg(3).ForeColorSel = &H80000005
    fg(3).BackColorSel = &H80&
    
    
    COLUMNAFCHPROD_ = 1
    COLUMNANUMPROD_ = 2
    COLUMNAPRODUCTO_ = 3
    COLUMNARECETA_ = 4
    COLUMNAUM_ = 5
    COLUMNACANTIDAD_ = 6
    COLUMNAENCARGADO_ = 7
    COLUMNALINEA_ = 8
    COLUMNAPORCENTAJEAPLICADO_ = 9
    COLUMNAHORINI_ = 10
    COLUMNAHORFIN_ = 11
    COLUMNAFCHFIN_ = 12
    COLUMNANUMOPE_ = 13
    COLUMNAPROCESADO_ = 14
    COLUMNAREPROC_ = 15
    COLUMNACERRADO_ = 16
    COLUMNANUMREGPROD_ = 17
    COLUMNAIDCRDET_ = 18
    COLUMNAIDRECETA_ = 19
    COLUMNAIDITEM_ = 20
    COLUMNAIDLINEA_ = 21
    COLUMNAIDRESP_ = 22
        
    GRID_COMBOLIST fg(3), COLUMNAPRODUCTO_
    GRID_COMBOLIST fg(3), COLUMNARECETA_
    GRID_COMBOLIST fg(3), COLUMNAENCARGADO_
    GRID_COMBOLIST fg(3), COLUMNALINEA_
        
    fg(3).ColEditMask(COLUMNAFCHPROD_) = "##/##/##"
    fg(3).ColEditMask(COLUMNAHORINI_) = "##:##"
    fg(3).Rows = fg(3).FixedRows
    
    fg(3).ColWidth(COLUMNAIDCRDET_) = 0
    fg(3).ColWidth(COLUMNAIDITEM_) = 0
    fg(3).ColWidth(COLUMNAIDLINEA_) = 0
    fg(3).ColWidth(COLUMNAIDRECETA_) = 0
    fg(3).ColWidth(COLUMNAIDRESP_) = 0
    fg(3).ColWidth(COLUMNAPROCESADO_) = 0
    fg(3).ColWidth(COLUMNANUMPROD_) = 0
    
'    fg(4).ColWidth(4) = 0
'    fg(4).ColWidth(5) = 0
    
    ELIMINARTODOS_ = False
    
    '************************
    ESTADO_ = 1
    '************************
End Sub

Private Sub cargarEstado(VALOR_ As Double)
    Dim A As Integer
    
    ' Se llena el estado
    For A = 0 To cbEstado.ListCount - 1
        If cbEstado.ItemData(A) = VALOR_ Then
            cbEstado.ListIndex = A
            Exit For
        End If
    Next A
End Sub

Private Sub llenarEstado(DETALLE_ As Boolean, Optional ESTADO_ As Double, _
                                        Optional ByRef FGGRID As VSFlexGrid, Optional COLUM_ As Integer)
'    Dim CAMPOS As String
'    Dim xRs As New ADODB.Recordset
'    Dim A As Integer
'
'    If DETALLE_ Then
'        For A = 0 To cbEstado.ListCount - 1
'            If cbEstado.ItemData(A) = ESTADO_ Then
'                cbEstado.ListIndex = A
'                Exit For
'            End If
'        Next A
'    Else
'        ' Se llena el flexGrid
'        cSQL = "SELECT * FROM mae_estados ORDER BY id"
'
'         cSQL = "SELECT * " _
'            + vbCr + "FROM mae_estados " _
'            + vbCr + "WHERE (((mae_estados.id) In (" & ESTADOPENDIENTE_ & "," & ESTADOPROCESADO_ & "," & ESTADOANULADO_ & "))) " _
'            + vbCr + "ORDER BY mae_estados.id;"
'
'        RST_Busq xRs, cSQL, xCon
'
'        If xRs.State = 0 Then Exit Sub
'        If xRs.RecordCount = 0 Then
'            MsgBox "No se ha encontrado estados, Ingrese estados", vbInformation, xTitulo
'            Exit Sub
'        End If
'
'        xRs.MoveFirst
'        CAMPOS = "#" & NulosN(xRs("id")) & ";" & UCase(NulosC(xRs("descripcion")))
'        xRs.MoveNext
'        While Not xRs.EOF
'            CAMPOS = CAMPOS & "|#" & NulosN(xRs("id")) & ";" & UCase(NulosC(xRs("descripcion")))
'            xRs.MoveNext
'        Wend
'        FGGRID.ColComboList(COLUM_) = CAMPOS
'
'        ' Se llena el comboBox
'        cbEstado.Clear
'        xRs.MoveFirst
'        While Not xRs.EOF
'            cbEstado.AddItem UCase(NulosC(xRs("descripcion")))
'            cbEstado.ItemData(cbEstado.NewIndex) = NulosN(xRs("id"))
'
'            xRs.MoveNext
'        Wend
'        cbEstado.ListIndex = 0
'    End If
End Sub

Private Function procesarLineaProduccion(ByRef FGRID_ As VSFlexGrid, Optional DISEÑO_ As Boolean = False) As Boolean
    Dim xRs As New ADODB.Recordset
    Dim CANTIDADAPROCESAR_ As Double
    Dim CANTIDAD_ As Double
    Dim IDLINEA_ As Double
    Dim IDCRDET_ As Double
    Dim IDITEM_ As Double
    Dim HORINI_ As String
    Dim FECHINI_ As Date
    Dim ESNUEVO_ As Boolean
            
    '*****************************
    Dim PORCENTAJEAPLICADO_ As Double
    '*****************************
            
    ' Se inicializan datos requeridos
    If DISEÑO_ Then
        If NulosC(fg(3).TextMatrix(fg(3).Row, COLUMNAFCHPROD_)) = "" Then ' Fecha de Inicio
            MsgBox "Ingrese Fecha de Programación", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            fg(3).Select fg(3).Row, COLUMNAFCHPROD_
            procesarLineaProduccion = False: Exit Function
        End If

        If NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDITEM_)) = 0 Then ' Producto
            MsgBox "Ingrese Producto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            fg(3).Select fg(3).Row, COLUMNAPRODUCTO_
            procesarLineaProduccion = False: Exit Function
        End If

        If NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNACANTIDAD_)) = 0 Then ' Cantidad
            MsgBox "Ingrese Cantidad", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            fg(3).Select fg(3).Row, COLUMNACANTIDAD_
            procesarLineaProduccion = False: Exit Function
        End If

        If fg(3).TextMatrix(fg(3).Row, COLUMNAIDRESP_) = "" Then ' Encargado
            MsgBox "Ingrese Encargado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            fg(3).Select fg(3).Row, COLUMNAENCARGADO_
            procesarLineaProduccion = False: Exit Function
        End If

        If fg(3).TextMatrix(fg(3).Row, COLUMNAHORINI_) = "" Then ' Hora de Inicio
            MsgBox "Ingrese Hora de Inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            fg(3).Select fg(3).Row, COLUMNAHORINI_
            procesarLineaProduccion = False: Exit Function
        End If
        
        IDLINEA_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDLINEA_))
        IDCRDET_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_))
        IDITEM_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDITEM_))
        HORINI_ = Format(fg(3).TextMatrix(fg(3).Row, COLUMNAHORINI_), "HH:mm")
        FECHINI_ = CDate(Format(fg(3).TextMatrix(fg(3).Row, COLUMNAFCHPROD_), "dd/mm/yyyy"))
        CANTIDAD_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNACANTIDAD_))
        
        '*************************************
        PORCENTAJEAPLICADO_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAPORCENTAJEAPLICADO_))
        '*************************************
    Else
        If NulosN(TxtMatProd.Text) = 0 Then ' Producto
            MsgBox "Ingrese Producto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            fg(3).Select fg(3).Row, COLUMNAPRODUCTO_
            procesarLineaProduccion = False: Exit Function
        End If

        If NulosN(TxtCant.Text) = 0 Then ' Cantidad
            MsgBox "Ingrese Cantidad", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            fg(3).Select fg(3).Row, COLUMNACANTIDAD_
            procesarLineaProduccion = False: Exit Function
        End If

        If NulosN(TxtIdEncarg.Text) = 0 Then ' Encargado
            MsgBox "Ingrese Encargado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            fg(3).Select fg(3).Row, COLUMNAENCARGADO_
            procesarLineaProduccion = False: Exit Function
        End If

        If Not IsDate(DTPHoras.Value) Then ' Hora de Inicio
            MsgBox "Ingrese Hora de Inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            fg(3).Select fg(3).Row, COLUMNAHORINI_
            procesarLineaProduccion = False: Exit Function
        End If
        IDLINEA_ = NulosN(TxtIdLineaDet.Text)
        IDCRDET_ = NulosN(LblIdCrDet.Caption)
        IDITEM_ = NulosN(TxtMatProd.Text)
        HORINI_ = Format(DTPHoras.Value, "HH:mm")
        FECHINI_ = CDate(Format(LblDia.Caption, "dd/mm/yyyy"))
        CANTIDAD_ = NulosN(TxtCant.Text)
        
        '*************************************
        PORCENTAJEAPLICADO_ = NulosN(txtEfic.Text)
        '*************************************
    End If
        
    ' Se filtra el registro involucrado
    RstTareasAux.Filter = "idcrdet = " & IDCRDET_ & ""
    
    ' Si no hay Tareas Procesadas anteriormente
    If RstTareasAux.RecordCount = 0 Then
        cSQL = "SELECT " & IDCRDET_ & " AS idcrdet, pro_receta.iditem, pro_lineadet.idtar, pro_lineadet.orden, pro_tareas.descripcion AS destar, pro_lineadet.factor, pro_lineadet.kghora AS costokg, pro_lineadet.numop AS numper, pro_lineadet.intervalo AS horarr, pro_lineadet.rdmto AS aplpor, '" & HORINI_ & "' AS horinitar, '" & FECHINI_ & "' AS fchini, -1 AS activo, pro_recetatar.idarea, pro_recetatar.idtiptrab, pro_recetatar.idformapag, mae_area.descripcion AS desarea, pro_tiptrab.descripcion AS destiptrab, pro_formapag.descripcion AS desformapag, pla_empleados.id AS idresp, pla_empleados.nombre AS nomresp " _
            + vbCr + "FROM ((((((((pro_lineadet LEFT JOIN pro_tareas ON pro_lineadet.idtar = pro_tareas.id) LEFT JOIN (pro_receta LEFT JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) ON pro_lineadet.idrec = pro_receta.id) LEFT JOIN pro_recetatar ON (pro_lineadet.idrec = pro_recetatar.idrec) AND (pro_lineadet.idtar = pro_recetatar.idtar)) LEFT JOIN pro_tiptrab ON pro_recetatar.idtiptrab = pro_tiptrab.id) LEFT JOIN pro_formapag ON pro_recetatar.idformapag = pro_formapag.id) LEFT JOIN mae_area ON pro_recetatar.idarea = mae_area.id) LEFT JOIN pro_area ON mae_area.id = pro_area.idarea) LEFT JOIN pro_emp ON pro_area.idper = pro_emp.id) LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id " _
            + vbCr + "Where (((pro_lineadet.idlineadet) = " & IDLINEA_ & ")) " _
            + vbCr + "ORDER BY pro_lineadet.orden;"
            
        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 0 Then procesarLineaProduccion = False: Exit Function
        
        If xRs.RecordCount = 0 Then
            MsgBox "No se encontro datos de la Linea de Produccion; Agregue una y procese de nuevo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            procesarLineaProduccion = False
            Exit Function
        End If
        
        ESNUEVO_ = True
    Else
        DEFINIR_RST_TMP xRs, RstTareasAux
        CARGAR_RST_TMP xRs, RstTareasAux
        
        ESNUEVO_ = False
    End If
    
    Dim RECORDSET_ As New ADODB.Recordset
    If ESNUEVO_ Then
        Set RECORDSET_ = xRs
    Else
        Set RECORDSET_ = RstTareasAux
    End If
    
    CANTIDADAPROCESAR_ = calcularRdmto(IDLINEA_, IDCRDET_, RECORDSET_, CANTIDAD_)
    procesarCronograma xRs, ESNUEVO_, CANTIDADAPROCESAR_, HORINI_, HORINI_, FECHINI_, IDITEM_, IDCRDET_, IDLINEA_, , , PORCENTAJEAPLICADO_
    
    ' Se carga las Tareas
    pCargarDatos FGRID_, False, True, False, False, False, DISEÑO_
    calcularDatosAdicionales DISEÑO_
    If frm(2).Visible Then Cmd(10).SetFocus
    procesarLineaProduccion = True
End Function

Private Sub aplicarCambios(Optional DISEÑO_ As Boolean = False)
    Dim FECHAINI_ As Date
    Dim FECHAFIN_ As Date
    Dim NUMPROG_ As String
    Dim IDCRDET_ As Double
    Dim IDITEM_ As Double
    Dim IDREC_ As Integer
    Dim RECETA_ As String
    Dim IDRESP_ As Integer
    Dim RESPONSABLE_ As String
    Dim IDLINEA_ As Integer
    Dim LINEA_ As String
    Dim CANTIDAD_ As Double
    Dim PRODUCTO_ As String
    Dim UM_ As String
    Dim NUMOPE_ As Double
    Dim CERRADO_ As Boolean
    Dim IDESTADO_ As Integer
    Dim IDREGPROD_ As Double
    Dim NUMPROD_ As String
    
    '*************************
    Dim PORCENTAJEAPLICADO_ As Double
    Dim REPROCESO_ As Double
    '*************************
        
    Dim NUMERODOCUMENTO_ As Integer
    Dim IDORD_ As Integer
    Dim FCHORD_ As String
    Dim NUMSER_ As String
    Dim NUMDOC_ As String
    'Dim IDRESP_ As Integer
    Dim IDTIPDOCREF_ As Integer
    Dim IDDOCREF_ As Integer
    'Dim IDREC_ As Integer
    Dim IDUNIMED_ As Integer
    'Dim CANTIDAD_ As Double
    'Dim IDLINEA_ As Integer
    Dim EFIC_ As Integer
    Dim HORINI_ As String
    Dim HORFIN_ As String
    Dim FCHFIN_ As String
    Dim NUMOP_ As Integer
    Dim REPROC_ As Boolean
    Dim IDESTADOPROC_ As Integer
    Dim xRs As New ADODB.Recordset
    Dim xRsTar As New ADODB.Recordset
    Dim xRsPer As New ADODB.Recordset
    Dim xRsRep As New ADODB.Recordset
        
    If DISEÑO_ Then
        NUMPROG_ = NulosC(fg(3).TextMatrix(fg(3).Row, COLUMNANUMPROD_))
        IDCRDET_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_))
        IDITEM_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDITEM_))
        IDREC_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDRECETA_))
        RECETA_ = NulosC(fg(3).TextMatrix(fg(3).Row, COLUMNARECETA_))
        IDRESP_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDRESP_))
        RESPONSABLE_ = NulosC(fg(3).TextMatrix(fg(3).Row, COLUMNAENCARGADO_))
        IDLINEA_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDLINEA_))
        LINEA_ = NulosC(fg(3).TextMatrix(fg(3).Row, COLUMNALINEA_))
        CANTIDAD_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNACANTIDAD_))
        PRODUCTO_ = NulosC(fg(3).TextMatrix(fg(3).Row, COLUMNAPRODUCTO_))
        UM_ = NulosC(fg(3).TextMatrix(fg(3).Row, COLUMNAUM_))
        '***************************
        PORCENTAJEAPLICADO_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAPORCENTAJEAPLICADO_))
        '***************************
    Else
        NUMPROG_ = NulosC(lblNumprod.Caption)
        IDCRDET_ = NulosN(LblIdCrDet.Caption)
        IDITEM_ = NulosN(TxtMatProd.Text)
        IDREC_ = NulosN(LblIdRec.Caption)
        RECETA_ = NulosC(TxtCodRec.Text)
        IDRESP_ = NulosN(TxtIdEncarg.Text)
        RESPONSABLE_ = NulosC(LblEncargado.Caption)
        IDLINEA_ = NulosN(TxtIdLineaDet.Text)
        LINEA_ = NulosC(LblLinea.Caption)
        CANTIDAD_ = NulosN(TxtCant.Text)
        PRODUCTO_ = NulosC(LblMatProd.Caption)
        UM_ = NulosC(LblUnidad.Caption)
        '***************************
        PORCENTAJEAPLICADO_ = NulosN(txtEfic.Text)
        '***************************
    End If

    NUMOPE_ = NulosN(lblntrabtot.Caption)
    'CERRADO_ = ckCerrado.Value
    
    '*************************************************
    IDESTADO_ = cbEstado.ItemData(cbEstado.ListIndex)
    NUMPROD_ = NulosC(lblnumregprod.Caption)
    IDREGPROD_ = NulosN(lblidprocorr.Caption)
    '*************************************************
    
    '*************************
    ' SE VERIFICAN LOS CAMPOS
    '*************************
    If NUMPROG_ = "" Then
        MsgBox "Ingrese un Numero de Produccion, para la programacion actual", _
                                                vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    ' Se actualiza el estado como estado actual
    ' Para Tareas
    RstTareas.Filter = "idcrdet = " & IDCRDET_
    RstTareasAux.Filter = "idcrdet = " & IDCRDET_
    limpiarRST RstTareas, False
    CARGAR_RST_TMP RstTareas, RstTareasAux
    ' Para personal
    RstPersonal.Filter = "idcrdet = " & IDCRDET_
    RstPersonalAux.Filter = "idcrdet = " & IDCRDET_
    limpiarRST RstPersonal, False
    CARGAR_RST_TMP RstPersonal, RstPersonalAux
    ' Para Reprocesos
    RstReproc.Filter = "idcrdet = " & IDCRDET_
    RstReprocAux.Filter = "idcrdet = " & IDCRDET_
    limpiarRST RstReproc, False
    CARGAR_RST_TMP RstReproc, RstReprocAux
    
    If RstReproc.RecordCount = 0 Then
        REPROCESO_ = 0
    Else
        REPROCESO_ = -1
    End If
    
    ' Los productos
    ' Se agrega o se modifica en el registro de Productos
    If DISEÑO_ Then
        FECHAINI_ = CDate(Format(NulosC(fg(3).TextMatrix(fg(3).Row, COLUMNAFCHPROD_)), "dd/mm/yyyy") & " " _
                                    & Format(NulosC(fg(3).TextMatrix(fg(3).Row, COLUMNAHORINI_)), "HH:mm"))
        
        FECHAFIN_ = CDate(Format(lblFchFin.Caption, "dd/mm/yyyy") & " " + Format(LblHorFin.Caption, "HH:mm"))
    Else
        FECHAINI_ = CDate(Format(LblDia.Caption, "dd/mm/yyyy") & " " + Format(DTPHoras.Value, "HH:mm"))
        FECHAFIN_ = CDate(Format(lblFchFin.Caption, "dd/mm/yyyy") & " " + Format(LblHorFin.Caption, "HH:mm"))
    End If
    
    If RstProductosAux.State = 0 Then DEFINIR_RST_TMP RstProductosAux, RstProductos
    
    limpiarRST RstProductosAux
    RstProductosAux.AddNew
    RstProductosAux("id") = IDCRDET_
    RstProductosAux("numprod") = 0 'NUMPROG_
    RstProductosAux("fchpro") = FECHAINI_
    RstProductosAux("fchfin") = FECHAFIN_
    RstProductosAux("horpro") = Format(FECHAINI_, "HH:mm")
    RstProductosAux("horfin") = Format(FECHAFIN_, "HH:mm")
    RstProductosAux("iditem") = IDITEM_
    RstProductosAux("idrec") = IDREC_
    RstProductosAux("codrec") = RECETA_
    RstProductosAux("idresp") = IDRESP_
    RstProductosAux("nomresp") = RESPONSABLE_
    RstProductosAux("idlinea") = IDLINEA_
    RstProductosAux("nomlinea") = LINEA_
    RstProductosAux("cantidad") = CANTIDAD_
    RstProductosAux("descripcion") = PRODUCTO_
    RstProductosAux("abrev") = UM_
    RstProductosAux("numop") = NUMOPE_
    RstProductosAux("estado") = IDESTADO_
    RstProductosAux("idprocorr") = IDREGPROD_
    RstProductosAux("efic") = PORCENTAJEAPLICADO_
    RstProductosAux("reproc") = REPROCESO_
    RstProductosAux.Update
    
    RstProductos.Filter = "id = " & IDCRDET_
    RstProductosAux.Filter = "id = " & IDCRDET_
    limpiarRST RstProductos, False
    CARGAR_RST_TMP RstProductos, RstProductosAux
    
    
    ' ------------------------------SE APLICAN LOS CAMBIOS CORRESPONDIENTES AL ESTADO
    Select Case IDESTADO_
        Case ESTADOPROCESADO_ ' Procesado
            ' Grabamos los registros relacionados
            'If Not grabarRelacionados(IDCRDET_, IDREGPROD_, NUMPROD_) Then Exit Sub
            '-----------------------------------------------------------------------
            '-------------------------------SE GRABA REGISTRO DE ORDEN DE PRODUCCION
            '-----------------------------------------------------------------------
            RstProductosAux.Filter = adFilterNone
            RstProductosAux.Filter = "id=" & IDCRDET_
            If RstProductosAux.RecordCount = 0 Then Exit Sub
            ' --------------------------------------------CABECERA
            IDORD_ = 0
            NUMSER_ = "0001"
            NUMERODOCUMENTO_ = HallaCodigoTabla("pro_ordenprod", xCon, "numdoc")
            NUMDOC_ = Format(NulosC(NUMERODOCUMENTO_), "0000000000")
            IDTIPDOCREF_ = 116
            IDDOCREF_ = IDCRDET_
            FCHORD_ = Format(RstProductosAux("fchpro"), "dd/mm/yyyy")
            IDRESP_ = NulosN(RstProductosAux("idresp"))
            IDREC_ = NulosC(RstProductosAux("idrec"))
            IDUNIMED_ = NulosN(Busca_Codigo(NulosN(RstProductosAux("iditem")), "id", "idunimed", "alm_inventario", "N", xCon))
            CANTIDAD_ = NulosN(RstProductosAux("cantidad"))
            IDLINEA_ = NulosN(RstProductosAux("idlinea"))
            EFIC_ = NulosN(RstProductosAux("efic"))
            HORINI_ = Format(NulosC(RstProductosAux("horpro")), "HH:mm")
            HORFIN_ = Format(NulosC(RstProductosAux("horfin")), "HH:mm")
            FCHFIN_ = Format(NulosC(RstProductosAux("fchfin")), "dd/mm/yyyy")
            NUMOP_ = NulosN(RstProductosAux("numop"))
            REPROC_ = NulosN(RstProductosAux("reproc"))
            IDESTADOPROC_ = ESTADOPENDIENTE_
            ' -------------------------------------------RECORDSET DE TAREAS
            RstTareasAux.Filter = adFilterNone
            RstTareasAux.Filter = "idcrdet=" & IDCRDET_
            If RstTareasAux.RecordCount > 0 Then RstTareasAux.MoveFirst
            If xRsTar.State = 0 Then
                cSQL = "SELECT TOP 1 * FROM pro_ordenprodtar;"
                Set xRs = Nothing
                RST_Busq xRs, cSQL, xCon
                DEFINIR_RST_TMP xRsTar, xRs
            End If
            limpiarRST xRsTar
            While Not RstTareasAux.EOF
                xRsTar.AddNew
                xRsTar("idtar") = NulosN(RstTareasAux.Fields("idtar"))
                xRsTar("cantsum") = NulosN(RstTareasAux.Fields("cantproc"))
                xRsTar("cantproc") = NulosN(RstTareasAux.Fields("cantproc"))
                xRsTar("numop") = NulosN(RstTareasAux.Fields("numper"))
                xRsTar("fchini") = RstTareasAux.Fields("fchini")
                xRsTar("fchfin") = RstTareasAux.Fields("fchfin")
                xRsTar("horini") = RstTareasAux.Fields("horinitar")
                xRsTar("horfin") = RstTareasAux.Fields("horfintar")
                xRsTar("durtar") = RstTareasAux.Fields("durtar")
                xRsTar("idsubresp") = RstTareasAux.Fields("idresp")
                xRsTar("idarea") = RstTareasAux.Fields("idarea")
                xRsTar("activo") = RstTareasAux.Fields("activo")
                xRsTar.Update
                RstTareasAux.MoveNext
            Wend
            ' ------------------------------------------RECORDSET DE PERSONAS
            RstPersonalAux.Filter = adFilterNone
            RstPersonalAux.Filter = "idcrdet=" & IDCRDET_
            If RstPersonalAux.RecordCount > 0 Then RstPersonalAux.MoveFirst
            If xRsPer.State = 0 Then
                cSQL = "SELECT TOP 1 * FROM pro_ordenprodpers;"
                Set xRs = Nothing
                RST_Busq xRs, cSQL, xCon
                DEFINIR_RST_TMP xRsPer, xRs
            End If
            limpiarRST xRsPer
            While Not RstPersonalAux.EOF
                xRsPer.AddNew
                xRsPer("idper") = NulosN(RstPersonalAux.Fields("idper"))
                xRsPer.Update
                RstPersonalAux.MoveNext
            Wend
            ' ------------------------------------------RECORDSET DE REPROCESOS
            RstReprocAux.Filter = adFilterNone
            RstReprocAux.Filter = "idcrdet=" & IDCRDET_
            If RstReprocAux.RecordCount > 0 Then RstReprocAux.MoveFirst
            If xRsRep.State = 0 Then
                cSQL = "SELECT TOP 1 * FROM pro_ordenprodreproc;"
                Set xRs = Nothing
                RST_Busq xRs, cSQL, xCon
                DEFINIR_RST_TMP xRsRep, xRs
            End If
            limpiarRST xRsRep
            While Not RstReprocAux.EOF
                xRsRep.AddNew
                xRsRep("idlotedet") = NulosN(RstReprocAux.Fields("idlotedet"))
                xRsRep("cantidad") = NulosN(RstReprocAux("cantidad"))
                xRsRep.Update
                RstReprocAux.MoveNext
            Wend
            
            If Not grabarOrdProd(FCHORD_, IDTIPDOCREF_, IDDOCREF_, IDRESP_, IDREC_, IDUNIMED_, CANTIDAD_, _
                IDLINEA_, EFIC_, HORINI_, HORFIN_, FCHFIN_, NUMOP_, REPROC_, NUMDOC_, _
                xRsTar, xRsPer, xRsRep, NUMSER_, IDORD_, IDESTADOPROC_, CInt(AnoTra), Month(Date), QueHace) Then
                
                MsgBox "Ocurrió un error al tratar de generar la Orden de Producción", _
                                                vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
        
        Case ESTADOANULADO_ ' Anulado
            ' Anulamos los registros Relacionados
            If Not cambiarEstadoRelacionados(IDREGPROD_, ESTADOANULADO_) Then Exit Sub
            
    End Select
    
    If DISEÑO_ Then
        fg(3).TextMatrix(fg(3).Row, COLUMNAHORFIN_) = Format(FECHAFIN_, FORMAT_HORA_SIN_SEGUNDO)
        fg(3).TextMatrix(fg(3).Row, COLUMNAFCHFIN_) = Format(FECHAFIN_, FORMAT_DATE)
        fg(3).TextMatrix(fg(3).Row, COLUMNANUMOPE_) = Format(NUMOPE_, "00")
        fg(3).TextMatrix(fg(3).Row, COLUMNACERRADO_) = IDESTADO_
        fg(3).TextMatrix(fg(3).Row, COLUMNANUMREGPROD_) = NUMPROD_
        fg(3).TextMatrix(fg(3).Row, COLUMNAREPROC_) = REPROCESO_
        
        If fg(3).TextMatrix(fg(3).Row, COLUMNANUMPROD_) = 0 Then
            ' Si se ha cambiado a un estado procesado
            If IDESTADO_ = ESTADOPROCESADO_ And ESTADO_ <> ESTADOPROCESADO_ Then
                fg(3).TextMatrix(fg(3).Row, COLUMNANUMPROD_) = -1
            Else
                ' Si se ha cambiado a un estado Pendiente
                If IDESTADO_ = ESTADOPENDIENTE_ And ESTADO_ <> ESTADOPENDIENTE_ Then
                    fg(3).TextMatrix(fg(3).Row, COLUMNANUMPROD_) = -1
                Else
                    fg(3).TextMatrix(fg(3).Row, COLUMNANUMPROD_) = 0
                End If
            End If
        End If
        
        If NulosC(fg(3).TextMatrix(fg(3).Row, COLUMNAHORFIN_)) <> "" _
                    And NulosC(fg(3).TextMatrix(fg(3).Row, COLUMNAFCHFIN_)) <> "" _
                    And NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNANUMOPE_)) <> 0 Then
                    
            fg(3).TextMatrix(fg(3).Row, COLUMNAPROCESADO_) = "PROCESADO"
            fg(3).Select fg(3).Row, 1, fg(3).Row, fg(3).Cols - 1
            fg(3).FillStyle = flexFillRepeat
            fg(3).CellBackColor = &H80000005
        End If
        
        If Grabar Then
            frm(2).Visible = False
        Else
            MsgBox "Ocurrió un error al tratar de aplicar los cambios solicitados", _
                                                vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
    Else
        ' Se Agrega en el calendario
        operaciones True, False, False, IDCRDET_
    End If
    
End Sub

Private Function verificarEstado(IDCRDET_ As Double) As Boolean
    Dim MENSAJE_ As String
    
    If Agregando Then Exit Function
    
    If cbEstado.ItemData(cbEstado.ListIndex) = 1 Then
        If ESTADO_ >= 2 Then
            MsgBox "No se puede pasar a un estado " & cbEstado.Text, vbInformation, xTitulo
            ' Se regresa a su estado anterior
            cargarEstado ESTADO_
            verificarEstado = False
            Exit Function
        End If
    End If
    
    verificarEstado = True
End Function

Private Sub cbEstado_Click()
    Dim Rpta As Integer
    Dim IDCRDET_ As Double
    Dim DISEÑO_ As Boolean
    Dim MENSAJE_ As String
    Dim NUMERODOCUMENTO_ As Integer
    
    If Agregando Then Exit Sub
    
    If CalCtrlCronog.Visible Then
        DISEÑO_ = False
    Else
        DISEÑO_ = True
    End If
    
    If DISEÑO_ Then
        IDCRDET_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_))
    Else
        IDCRDET_ = NulosN(LblIdCrDet.Caption)
    End If
            
    Select Case cbEstado.ItemData(cbEstado.ListIndex)
        Case ESTADOPENDIENTE_ ' Pendiente
            If ESTADO_ > ESTADOPENDIENTE_ Then
                MsgBox "No se puede cambiar el estado a " & cbEstado.Text, vbInformation, xTitulo
                llenarEstado True, ESTADO_
            End If
            Exit Sub
        
        Case ESTADOPROCESADO_ ' Procesado
            If ESTADO_ < ESTADOPROCESADO_ Then
                Rpta = MsgBox("Cambiar el estado a " & cbEstado.Text & " hara que el sistema genere automaticamente " _
                                    & "los registros relacionados en las areas afines;" _
                                    + vbCr + "¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)

                If Rpta = vbNo Then
                    Agregando = True
                    cargarEstado ESTADO_
                    Agregando = False
                End If
            Else
                MsgBox "No se puede pasar a un estado " & cbEstado.Text, vbInformation, xTitulo
                Agregando = True
                cargarEstado ESTADO_
                Agregando = False
            End If
            Exit Sub
        
        Case ESTADOANULADO_ ' Anulada
            If ESTADO_ = ESTADOPROCESADO_ Then
                If Not verificarCambioEstado(NulosN(lblidprocorr.Caption), MENSAJE_) Then
                    MsgBox "No se puede pasar a un estado " & cbEstado.Text, vbInformation, xTitulo
                    Agregando = True
                    cargarEstado ESTADO_
                    Agregando = False
                End If
            Else
                MsgBox "No se puede cambiar el estado a " & cbEstado.Text, vbInformation, xTitulo
                Agregando = True
                cargarEstado ESTADO_
                Agregando = False
            End If
            Exit Sub
            
    End Select
End Sub

Private Sub cbFecha_Click()
    'frm(2).Visible = False
    llenarDatos True, , cbfecha.Text
End Sub

Private Sub cbFecha_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbsemcamb_Click()
    ' Se carga el codigo del idcrdet de la semana seleccionada
    LblDetProd(1).Caption = Busca_Codigo(NulosN(cbsemcamb.Text), "semana", "id", "pro_cronograma", "N", xCon)
    ' Se carga los dias correspondientes a la semana
    cargarDiasSemanaReg NulosN(cbsemcamb.Text)
End Sub

Private Function cambiarEstadoRelacionados(IDREGPROD_ As Double, ESTADO_ As Double) As Boolean
    Dim ID_ As Double
    
    On Error GoTo ERROR_
    
    ' Ordenes de Produccion
    cSQL = "UPDATE pro_producciondet SET pro_producciondet.estado = " & ESTADO_ & " " _
        + vbCr + "WHERE (((pro_producciondet.idtipdocref)=116) AND ((pro_producciondet.iddocref)=" & IDREGPROD_ & "));"
    
    xCon.Execute cSQL
    
    ' Solicitud de Materiales
    cSQL = "UPDATE pro_ordenprod SET pro_ordenprod.estado = " & ESTADO_ & " " _
        + vbCr + "WHERE (((pro_ordenprod.idprocorr)=" & IDREGPROD_ & "));"
    
    xCon.Execute cSQL
    
    ' Registros de Planillas
    cSQL = "UPDATE pro_controltardet SET pro_controltardet.estado = " & ESTADO_ & " " _
        + vbCr + "WHERE (((pro_controltardet.idprocorr)=" & IDREGPROD_ & "));"
    
    xCon.Execute cSQL
    
    ' Salidas de Almacen
    cSQL = "UPDATE alm_ingreso SET alm_ingreso.estado = " & ESTADO_ & " " _
        + vbCr + "WHERE (((alm_ingreso.idprocorr)=" & IDREGPROD_ & "));"

    xCon.Execute cSQL
            
    ' GRABAMOS LOS MOVIMIENTOS
    ' REGISTRO DE PRODUCCION
    ID_ = Busca_Codigo(IDREGPROD_, "corr", "idpro", "pro_producciondet", "N", xCon)
    GrabarOperacion xIdUsuario, 92, 7, xHorIni, Time, Date, xCon, ID_
    ' SOLICITUD DE MATERIALES
    ID_ = Busca_Codigo(IDREGPROD_, "idprocorr", "idord", "pro_ordenproddet", "N", xCon)
    GrabarOperacion xIdUsuario, 54, 7, xHorIni, Time, Date, xCon, ID_
    ' REGISTRO DE PLANILLA
    ID_ = Busca_Codigo(IDREGPROD_, "idprocorr", "idctr", "pro_controltardet", "N", xCon)
    GrabarOperacion xIdUsuario, 179, 7, xHorIni, Time, Date, xCon, ID_
    ' INGRESOSOS Y SALIDAS DE ALMACEN
    ID_ = Busca_Codigo(IDREGPROD_, "idprocorr", "id", "alm_ingreso", "N", xCon)
    GrabarOperacion xIdUsuario, 8, 7, xHorIni, Time, Date, xCon, ID_
                
    cambiarEstadoRelacionados = True
    Exit Function
ERROR_:
    MsgBox "Ha ocurrido un error al tratar de cambiar de estado", vbInformation, xTitulo
    cambiarEstadoRelacionados = False
End Function

Private Function grabarRelacionados(IDCRDET_ As Double, ByRef IDREGPROD_ As Double, ByRef NUMPROD_ As String) As Boolean
    Dim Rpta As Integer
    Dim NINGUNERROR_ As Boolean
    Dim MENSAJE_ As String
    Dim DISEÑO_ As Boolean
    'Dim IDREGPROD_ As Double    ' Id de Produccion Generado
    Dim IDREGSOL_ As Double     ' Id de Solicitud de Materiales Generado
    'Dim NUMPROD_ As String      ' Numero de Produccion
    Dim NUMSOL_ As String       ' Numero de Solicitud de Materiales Generado
    Dim ID_ As Double
    
    DISEÑO_ = Not CalCtrlCronog.Visible
            
    NINGUNERROR_ = True
    
    If DISEÑO_ Then
        If NulosC(fg(3).TextMatrix(fg(3).Row, COLUMNAPROCESADO_)) = "" Then
            MENSAJE_ = "El Evento actual no esta procesado y no se puede aprobar"
            NINGUNERROR_ = False
        End If
    End If
    
    CentrarFrm frm(3)
    frm(3).Visible = True
    
    xCon.BeginTrans
    
    If NINGUNERROR_ Then
        ' Se verifica la existencia del registro
        RstProductos.Filter = "id = " & IDCRDET_
        If RstProductos.RecordCount = 0 Then
            MENSAJE_ = "No se puede realizar esta operación para un registro nuevo"
            NINGUNERROR_ = False
        End If
    End If

    'RstProductos("cerrado") = True

    LblProg.Caption = "REGISTRO DE PRODUCCION"
    frm(3).Refresh
    If NINGUNERROR_ Then
        MENSAJE_ = "Ha ocurrido un error al intentar crear el Registro de Producción; se cancelara la operación"
        NINGUNERROR_ = GrabarProduccion(IDCRDET_, IDREGPROD_)
    End If

    LblProg.Caption = "SOLICITUD DE MATERIALES"
    frm(3).Refresh
    If NINGUNERROR_ Then
        MENSAJE_ = "Ha ocurrido un error al intentar crear la Solicitud de Materiales; se cancelara la operación"
        NINGUNERROR_ = grabarSolicitud(IDCRDET_, IDREGSOL_, IDREGPROD_)
    End If

    LblProg.Caption = "REGISTRO DE PLANILLA"
    frm(3).Refresh
    If NINGUNERROR_ Then
        MENSAJE_ = "Ha ocurrido un error al intentar crear el Registro de Planilla; se cancelara la operación"
        NINGUNERROR_ = GrabarPlanilla(IDCRDET_, IDREGPROD_)
    End If

    LblProg.Caption = "REGISTRO DE ALMACEN"
    frm(3).Refresh
    If NINGUNERROR_ Then
        MENSAJE_ = "Ha ocurrido un error al intentar crear el Registro de Ingreso de Produccion; se cancelara la operación"
        NINGUNERROR_ = GrabarAlmacen(IDCRDET_, IDREGSOL_, IDREGPROD_)
    End If

    LblProg.Caption = "APLICANDO CAMBIOS"
    frm(3).Refresh
    If NINGUNERROR_ Then
        xCon.CommitTrans
    
        ' GRABAMOS LOS MOVIMIENTOS
        ' REGISTRO DE PRODUCCION
        ID_ = Busca_Codigo(IDREGPROD_, "corr", "idpro", "pro_producciondet", "N", xCon)
        GrabarOperacion xIdUsuario, 92, 6, xHorIni, Time, Date, xCon, ID_
        ' SOLICITUD DE MATERIALES
        ID_ = Busca_Codigo(IDREGPROD_, "idprocorr", "idord", "pro_ordenproddet", "N", xCon)
        GrabarOperacion xIdUsuario, 54, 6, xHorIni, Time, Date, xCon, ID_
        ' REGISTRO DE PLANILLA
        ID_ = Busca_Codigo(IDREGPROD_, "idprocorr", "idctr", "pro_controltardet", "N", xCon)
        GrabarOperacion xIdUsuario, 179, 6, xHorIni, Time, Date, xCon, ID_
        ' INGRESSOS Y SALIDAS DE ALMACEN
        ID_ = Busca_Codigo(IDREGSOL_, "idorddet", "id", "alm_ingreso", "N", xCon)
        GrabarOperacion xIdUsuario, 8, 6, xHorIni, Time, Date, xCon, ID_
        
        Agregando = True
        NUMPROD_ = Busca_Codigo(IDREGPROD_, "corr", "numparte", "pro_producciondet", "N", xCon)
        
        Agregando = False
        grabarRelacionados = True
    Else
        xCon.RollbackTrans
        IDREGPROD_ = 0
        NUMPROD_ = ""
        MsgBox MENSAJE_, vbInformation, xTitulo
        grabarRelacionados = False
    End If

    frm(3).Visible = False

End Function

Private Sub cmd_Click(Index As Integer)
    Dim xFrm As New sgi2_produccion.produccion
    Dim DISEÑO_ As Boolean
    Dim xCampos(2, 4) As String
    Dim xRs As New ADODB.Recordset
    Dim nTitulo As String
    Dim Rpta As Integer
    
    DISEÑO_ = Not CalCtrlCronog.Visible
    Select Case Index
        Case 0 ' Agregar Producto
            agregarCampos True, False
            
        Case 1 ' Establecer propiedades de procesado
            aplicarPropiedades False, True
            CentrarFrm frm(0)
            frm(0).ZOrder 0
            frm(0).Visible = True
            
        Case 2 ' Procesar la linea
            procesarLineaProduccion fg(0), DISEÑO_
            
        Case 3 ' Agregar tarea
            agregarCampos False, True
            ' Se carga al personal relacionado con esa tarea si es que lo hubiera
            RstPersonalAux.Filter = "idcrdet = " & NulosN(LblIdCrDet.Caption) & _
                                " And idtar = " & NulosN(fg(0).TextMatrix(fg(0).Row, 11)) & ""
            pCargarDatos fg(1), True, False, , , False
            
        Case 4 ' Agregar Personal
            procesarPersonal True, False, False, False, DISEÑO_
            
        Case 5 ' Listar personal
            procesarPersonal False, True, False, False, DISEÑO_
            
        Case 6 ' Eliminar Personal
            procesarPersonal False, False, True, False, DISEÑO_
            
        Case 7 ' Eliminar Todos
            procesarPersonal False, False, False, True, DISEÑO_
            
        Case 8 ' Ver Ranking
            LbNumSel.Caption = 0
            OptSel(1).Value = True
            ' Se procesa el ranking para mostrarlo
            procesarRanking
            
        Case 9 'Cancelar Migrar
            cbfchcamb.Clear
            cbsemcamb.Clear
            frm(1).Visible = False
        
        Case 10 ' Acepta Agregar/Modificar Detalle
            Dim FILASEL_ As Integer
            If DISEÑO_ Then
                FILASEL_ = fg(3).Row
            End If
            
            aplicarCambios DISEÑO_
            If DISEÑO_ Then
                fg(3).Select FILASEL_, COLUMNAHORINI_
            End If
            
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
            
            frm(2).Visible = False
            
        Case 12 ' Aceptar Propiedades de procesado
            aplicarPropiedades True
            frm(0).Visible = False
            
        Case 13 ' Cancela Propiedades de procesado
            frm(0).Visible = False
            
        Case 14 ' Adicionar de Ranking
            procesarRanking False, True
            
        Case 15 ' Cancela Procesar Ranking
            frm(4).Visible = False
        
        Case 16 ' Elegir Receta
            agregarCampos False, False, False, True
            
        Case 17 ' Aceptar migrar evento
            Rpta = MsgBox("El registro asociado debe de estar previamente grabado para que los cambios surtan efecto; desea continuar el cambio no podra deshacerse?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbNo Then Exit Sub
            xCon.BeginTrans
            ' Se actualiza pro_cronogramadet
            cSQL = "UPDATE pro_cronogramadet SET pro_cronogramadet.idcr = " & NulosN(LblDetProd(1).Caption) & ", pro_cronogramadet.fchpro = CDate('" & NulosC(cbfchcamb.Text) & "'), pro_cronogramadet.fchfin = CDate('" & NulosC(CDate(cbfchcamb.Text) + NulosN(LblDetProd(2))) & "')" _
                    + vbCr + "WHERE (((pro_cronogramadet.id)=" & NulosN(LblDetProd(0).Caption) & "));"
            xCon.Execute cSQL
            ' Se actualiza pro_cronogramapers
            cSQL = "UPDATE pro_cronogramapers SET pro_cronogramapers.idcr = " & NulosN(LblDetProd(1).Caption) & " " _
                    + vbCr + "WHERE (((pro_cronogramapers.idcrdet)=" & NulosN(LblDetProd(0).Caption) & "));"
            xCon.Execute cSQL
            ' Se actualiza pro_cronogramatarea
            cSQL = "UPDATE pro_cronogramatarea SET pro_cronogramatarea.idcr = " & NulosN(LblDetProd(1).Caption) & ", pro_cronogramatarea.fchpro = CDate('" & NulosC(cbfchcamb.Text) & "'), pro_cronogramatarea.fchini = CDate('" & NulosC(cbfchcamb.Text) & "'), pro_cronogramatarea.fchfin = CDate('" & NulosC(CDate(cbfchcamb.Text) + NulosN(LblDetProd(2))) & "')" _
                    + vbCr + "WHERE (((pro_cronogramatarea.idcrdet)=" & NulosN(LblDetProd(0).Caption) & "));"
            xCon.Execute cSQL
            
            xCon.CommitTrans
            
            RstProductos.Filter = "id = " & NulosN(LblDetProd(0).Caption)
            RstTareas.Filter = "idcrdet = " & NulosN(LblDetProd(0).Caption)
            RstPersonal.Filter = "idcrdet = " & NulosN(LblDetProd(0).Caption)
            
            limpiarRST RstProductos, False
            limpiarRST RstTareas, False
            limpiarRST RstPersonal, False
            
            llenarDatos DISEÑO_
            
            cbsemcamb.Clear
            cbfchcamb.Clear
            frm(1).Visible = False
            
        Case 18 ' Escoger Encargado de Linea
            agregarCampos False, False, True
        
        Case 19 ' Imprimir Linea
            Dim IDCRDET_ As Double
            If DISEÑO_ Then
                IDCRDET_ = fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_)
            Else
                IDCRDET_ = NulosN(LblIdCrDet.Caption)
            End If
            
            RstProductos.Filter = adFilterNone
            RstProductos.Filter = "id = " & IDCRDET_
            
            If RstProductos.State = 0 Then Exit Sub
            If RstProductos.RecordCount = 0 Then Exit Sub
            
            If RstProductos("estado") <> ESTADOPROCESADO_ Then
                MsgBox "Un registro en este estado no se puede Imprimir", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
            
            RstProductos.Filter = adFilterNone
            With FrmVsPrinter.Vs
                .StartDoc
                Me.MousePointer = vbHourglass
                ImprimirLinea IDCRDET_
                Me.MousePointer = vbDefault
                .EndDoc
            End With
            'Muestra la preimagen de la impresion
            FrmVsPrinter.WindowState = 2
            FrmVsPrinter.Show
            
        Case 20 ' Buscar Linea
            agregarCampos False, False, False, False, True
        
        Case 21 'Seleccionar Personal Vista Diseño
            procesarPersonal False, True, False, False, True
        
        Case 22 ' Agregar Reproceso
            procesarReproceso True, False, False, False
            
        Case 23 ' Eliminar Reproceso
            procesarReproceso False, False, True, False
        
        Case 24 ' Seleccionar Reproceso
            procesarReproceso False, True, False, False
        
        Case 25 ' Eliminar Todos Reproceso
            procesarReproceso False, False, False, True
        
    End Select
End Sub

Private Sub procesarRanking(Optional MOSTRAR_ As Boolean = True, _
                                Optional AGREGAR_ As Boolean = False)
    Dim RstRanking As New ADODB.Recordset
    Dim A As Integer
    Dim nSQLId_0 As String
    Dim nSQLId_1 As String
    Dim nSQLId_2 As String
    Dim FECHA_ As String
    Dim REINTENTO_ As Boolean
    Dim IDRECETA_ As Double
    Dim IDTAREA_ As Double
    Dim PRODUCTO_ As String
    Dim TAREA_ As String
    Dim IDCRDET_ As Double
    Dim DISEÑO_ As Boolean
    
    If CalCtrlCronog.Visible Then DISEÑO_ = False Else DISEÑO_ = True
    
    If DISEÑO_ Then
        IDCRDET_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_))
        IDRECETA_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDRECETA_))
        IDTAREA_ = NulosN(fg(0).TextMatrix(fg(0).Row, 11))
        PRODUCTO_ = NulosC(fg(3).TextMatrix(fg(3).Row, COLUMNAPRODUCTO_))
        TAREA_ = NulosC(fg(0).TextMatrix(fg(0).Row, 2))
        FECHA_ = NulosC(fg(3).TextMatrix(fg(3).Row, COLUMNAFCHPROD_))
    Else
        IDCRDET_ = NulosN(LblIdCrDet.Caption)
        IDRECETA_ = NulosN(LblIdRec.Caption)
        IDTAREA_ = NulosN(fg(0).TextMatrix(fg(0).Row, 11))
        PRODUCTO_ = NulosC(LblMatProd.Caption)
        TAREA_ = NulosC(fg(0).TextMatrix(fg(0).Row, 2))
        FECHA_ = NulosC(LblDia.Caption)
    End If
        
    If MOSTRAR_ Then
On Error GoTo ERROR_AL_MOSTRAR
        
        LblProd2.Caption = PRODUCTO_
        LblTarea2.Caption = TAREA_
        
        ' Generar la lista de personal para no considerar en la lista
        RstPersonalAux.Filter = "idcrdet = " & IDCRDET_ & ""
        nSQLId_0 = GENERAR_SQL_ID_RST(RstPersonalAux, "idper", " AND pro_controltardet.idref", "NOT IN", True)
        nSQLId_2 = GENERAR_SQL_ID_RST(RstPersonalAux, "idper", " AND pro_controltardetgr.idper", "NOT IN", True)
        
        REINTENTO_ = False
REINTENTAR:
        nSQLId_1 = GENERAR_SQL_ID_RST(buscarAsistencia(FECHA_), "idemp", " AND pla_empleados.id", "IN", True)
        
        ' Si no hay datos de Asistencia se busca un dia antes
        If nSQLId_1 = "" Then
            If Not REINTENTO_ Then
                REINTENTO_ = True
                FECHA_ = Format(CDate(FECHA_) - 1, "dd/mm/yyyy")
                GoTo REINTENTAR
            End If
            MsgBox "No se encontro datos de la Asistencia; Se mostrara a todo el Personal", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
        
'        cSQL = "SELECT pro_controltardet.tipo, pro_controltardet.idref AS idper, pla_empleados.nombre, pro_receta.iditem, alm_inventario.descripcion AS producto, pro_controltardet.idtar, pro_tareas.abrev AS tarea, Sum(pro_controltardet.cant) AS SumaDecant, Last(pro_controltar.fchtra) AS ÚltimoDefchtra, Sum(1) AS diasTrab, pla_empleados.numdoc " _
'            + vbCr + "FROM pro_controltar INNER JOIN (pro_tareas RIGHT JOIN (mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN ((pro_receta RIGHT JOIN pro_controltardet ON pro_receta.id = pro_controltardet.idrec) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON alm_inventario.id = pro_receta.iditem) ON mae_unidades.id = pro_controltardet.idunimed) ON pro_tareas.id = pro_controltardet.idtar) ON pro_controltar.id = pro_controltardet.idctr " _
'            + vbCr + "GROUP BY pro_controltardet.tipo, pro_controltardet.idref, pla_empleados.nombre, pro_receta.iditem, alm_inventario.descripcion, pro_controltardet.idtar, pro_tareas.abrev, alm_inventario.descripcion, pro_tareas.abrev, pla_empleados.numdoc " _
'            + vbCr + "Having (((pro_controltardet.Tipo) = 1) And ((pla_empleados.nombre) Is Not Null) And ((pro_receta.iditem) = " & NulosN(TxtMatProd.Text) & ") And ((pro_controltardet.idtar) = " & IDTAREA_ & ") And ((pro_tareas.abrev) Is Not Null)) " & nSQLId_0 & nSQLId_1 _
'            + vbCr + "ORDER BY alm_inventario.descripcion, pro_tareas.abrev; " _
'            + vbCr + "Union " _
'            + vbCr + "SELECT pro_controltardet.tipo, pro_controltardetgr.idper, pla_empleados.nombre, pro_receta.iditem, alm_inventario.descripcion AS producto, pro_controltardet.idtar, pro_tareas.abrev AS tarea, Sum(pro_controltardetgr.cant) AS SumaDecant, Last(pro_controltar.fchtra) AS ÚltimoDefchtra, Sum(1) AS diasTrab, pla_empleados.numdoc " _
'            + vbCr + "FROM pro_controltar INNER JOIN (((pro_tareas RIGHT JOIN (mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN (pro_receta RIGHT JOIN pro_controltardet ON pro_receta.id = pro_controltardet.idrec) ON alm_inventario.id = pro_receta.iditem) ON mae_unidades.id = pro_controltardet.idunimed) ON pro_tareas.id = pro_controltardet.idtar) LEFT JOIN pro_controltardetgr ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id) ON pro_controltar.id = pro_controltardet.idctr " _
'            + vbCr + "GROUP BY pro_controltardet.tipo, pro_controltardetgr.idper, pla_empleados.nombre, pro_receta.iditem, alm_inventario.descripcion, pro_controltardet.idtar, pro_tareas.abrev, alm_inventario.descripcion, pro_tareas.abrev, pla_empleados.numdoc " _
'            + vbCr + "HAVING (((pro_controltardet.tipo)=2) AND ((pla_empleados.nombre) Is Not Null) AND ((pro_receta.iditem)= " & NulosN(TxtMatProd.Text) & ") AND ((pro_controltardet.idtar)= " & IDTAREA_ & ") AND ((pro_tareas.abrev) Is Not Null)) " & nSQLId_0 & nSQLId_1 _

        ' Se busca en tareas individuales
        cSQL = "SELECT pro_controltardet.tipo, pro_controltardet.idref AS idemp, pla_empleados.nombre, pla_empleados.numdoc, Sum(pro_controltardet.cant) AS totalcant, Sum(pro_controltardet.tothor) AS totalhor " _
            + vbCr + "FROM pro_controltar RIGHT JOIN (pro_tareas RIGHT JOIN (mae_unidades RIGHT JOIN (pro_controltardet LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON mae_unidades.id = pro_controltardet.idunimed) ON pro_tareas.id = pro_controltardet.idtar) ON pro_controltar.id = pro_controltardet.idctr " _
            + vbCr + "GROUP BY pro_controltardet.tipo, pro_controltardet.idref, pla_empleados.nombre, pla_empleados.numdoc, pro_controltardet.idtar, pla_empleados.id, pro_controltardet.idrec " _
            + vbCr + "HAVING (((pro_controltardet.Tipo) = 1) And ((pro_controltardet.idtar)=" & IDTAREA_ & ") AND ((pro_controltardet.idrec)=" & IDRECETA_ & ")) " & nSQLId_0 & nSQLId_1 _
                    
        RST_Busq RstRanking, cSQL, xCon
        
        fg(2).Rows = 1
        If RstRanking.State = 0 Then Exit Sub
        
        If RstRanking.RecordCount = 0 Then
            ' Se busca en tareas Grupales
            cSQL = "SELECT pro_controltardet.tipo, pro_controltardetgr.idper AS idemp, pla_empleados.nombre, pla_empleados.numdoc, Sum(pro_controltardetgr.cant) AS totalcant, Sum(pro_controltardetgr.tothor) AS totalhor " _
                + vbCr + "FROM pro_controltar INNER JOIN (((pro_tareas RIGHT JOIN pro_controltardet ON pro_tareas.id = pro_controltardet.idtar) LEFT JOIN pro_controltardetgr ON (pro_controltardet.idctr = pro_controltardetgr.idctr) AND (pro_controltardet.corr = pro_controltardetgr.corr)) LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id) ON pro_controltar.id = pro_controltardet.idctr " _
                + vbCr + "GROUP BY pro_controltardet.tipo, pro_controltardetgr.idper, pla_empleados.nombre, pla_empleados.numdoc, pro_controltardet.idtar, pla_empleados.id, pro_controltardet.idrec " _
                + vbCr + "HAVING (((pro_controltardet.tipo)=2) AND ((pro_controltardet.idtar)=" & IDTAREA_ & ") AND ((pro_controltardet.idrec)=" & IDRECETA_ & ")) " & nSQLId_2 & nSQLId_1 _
            
            Set RstRanking = Nothing
            RST_Busq RstRanking, cSQL, xCon
            
            If RstRanking.RecordCount = 0 Then
                MsgBox "No se ha encontrado registros que coincidan con la busqueda en los tipos de trabajo :" _
                + vbCr + "Grupal , Individual", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                GoTo SALIR_
            End If
        End If
        
        ' Se llenan los Datos
        RstRanking.MoveFirst
        For A = 1 To RstRanking.RecordCount
            fg(2).Rows = fg(2).Rows + 1
            fg(2).TextMatrix(A, 1) = 0
            fg(2).TextMatrix(A, 2) = A
            fg(2).TextMatrix(A, 3) = RstRanking("numdoc")
            fg(2).TextMatrix(A, 4) = RstRanking("nombre")
            fg(2).TextMatrix(A, 7) = RstRanking("idemp")
            fg(2).TextMatrix(A, 8) = Format(RstRanking("totalcant"), "0.00")
            fg(2).TextMatrix(A, 9) = Format(RstRanking("totalhor"), "0.00")
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
        
SALIR_:
        CentrarFrm frm(4)
        ' Se pone en primer plano
        frm(4).ZOrder 0
        frm(4).Visible = True
        Exit Sub
ERROR_AL_MOSTRAR:
        MsgBox "Ocurrio un Error al Visualizar, verifique que el Servidor este activo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    If AGREGAR_ Then
        Dim contador As Integer
        Dim num As Double
        
        num = NulosN(lblntrab.Caption) - (fg(1).Rows - 1)
        
        If LIMITARNUMEROPERSONAL_ Then
            If num <= 0 Then
                MsgBox "La Tarea actual no puede admitir mas Personal", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
        End If
        
        For A = 1 To fg(2).Rows - 1
            If LIMITARNUMEROPERSONAL_ Then
                If num <= 0 Then
                    MsgBox "La Tarea actual no puede admitir mas Personal, solo se agregara al personal necesario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    Exit For
                End If
            End If
            
            If fg(2).TextMatrix(A, 1) = 0 Then GoTo SIGUIENTE
            ' agregando los datos al rst temporal
            RstPersonalAux.AddNew
            RstPersonalAux("idcrdet") = IDCRDET_
            RstPersonalAux("idtar") = IDTAREA_
            RstPersonalAux("idper") = fg(2).TextMatrix(A, 7)
            RstPersonalAux("numdoc") = fg(2).TextMatrix(A, 3)
            RstPersonalAux("nombre") = fg(2).TextMatrix(A, 4)
            RstPersonalAux("activo") = True
            num = num - 1
SIGUIENTE:
        Next A
        RstPersonalAux.Filter = adFilterNone
        pCargarDatos fg(1), , , , , False, DISEÑO_
        frm(4).Visible = False
        totalizarPersonal
    End If
End Sub

Private Function buscarAsistencia(FECHA_CONSULTA As String) As ADODB.Recordset
    ' El recordset para acceder a los datos
    Dim RstAsistencia As ADODB.Recordset
    
    cSQL = "SELECT pla_recmarcacion.idemp, pla_recmarcacion.dia " _
        + vbCr + "From pla_recmarcacion " _
        + vbCr + "GROUP BY pla_recmarcacion.idemp, pla_recmarcacion.dia " _
        + vbCr + "HAVING (((pla_recmarcacion.dia)=CDate('" & FECHA_CONSULTA & "'))) " _
        + vbCr + "ORDER BY pla_recmarcacion.idemp;"
    
    Set RstAsistencia = New ADODB.Recordset
    ' Abrir el recordset de forma estática, no vamos a cambiar datos
    RST_Busq RstAsistencia, cSQL, xCon
    
    Set buscarAsistencia = RstAsistencia
    
'
'    ' Datos para la consulta
'    Dim CONS_FECH_ASISTENCIA As String
'
'    ' CONSULTA DE FECHA DE ASISTENCIA
'    CONS_FECH_ASISTENCIA = "(TEMPUS.MARCACIONES.FECHA = CAST('" & FECHA_CONSULTA & "' AS datetime)) "
'
'
'    ' CONSULTA
'    cSQL = "SELECT TEMPUS.EMPRESAS.NOMBRE AS EMP, " _
'                    + vbCr + "TEMPUS.PERSONAL.APELLIDO_PATERNO + ' ' + TEMPUS.PERSONAL.APELLIDO_MATERNO + ' ' + TEMPUS.PERSONAL.NOMBRES AS NOMPER, " _
'                    + vbCr + "CONVERT(varchar(12), TEMPUS.PERSONAL.FECHA_DE_NACIMIENTO, 103) AS FECHNAC, CONVERT(varchar(12), " _
'                    + vbCr + "TEMPUS.PERSONAL.FECHA_DE_INGRESO, 103) AS FECHING, TEMPUS.PERSONAL.DNI, CONVERT(varchar(12), TEMPUS.MARCACIONES.FECHA, 103) AS FECHMARC, " _
'                    + vbCr + "CONVERT(varchar(10), TEMPUS.MARCACIONES.HORA, 108) AS HORMARC, TEMPUS.CARGOS.DESCRIPCION " _
'            + vbCr + "FROM TEMPUS.MARCACIONES INNER JOIN " _
'                    + vbCr + "TEMPUS.PERSONAL ON TEMPUS.MARCACIONES.CODIGO = TEMPUS.PERSONAL.CODIGO AND " _
'                    + vbCr + "TEMPUS.MARCACIONES.EMPRESA = TEMPUS.PERSONAL.EMPRESA INNER JOIN " _
'                    + vbCr + "TEMPUS.EMPRESAS ON TEMPUS.PERSONAL.EMPRESA = TEMPUS.EMPRESAS.EMPRESA INNER JOIN " _
'                    + vbCr + "TEMPUS.CARGOS ON TEMPUS.PERSONAL.CARGO = TEMPUS.CARGOS.CARGO " _
'            + vbCr + "WHERE " & CONS_FECH_ASISTENCIA & " " _
'            + vbCr + "ORDER BY TEMPUS.MARCACIONES.FECHA, TEMPUS.PERSONAL.APELLIDO_PATERNO"
'
'    ' Abrir el recordset de forma estática, no vamos a cambiar datos
'    RST_Busq RstAsistencia, cSQL, con_SQLS
'
'    Set buscarAsistencia = RstAsistencia
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

Private Sub procesarReproceso(AGREGAR_ As Boolean, LISTAR_ As Boolean, _
                                    ELIMINAR_ As Boolean, ELIMTODOS_ As Boolean)
    If QueHace = 3 Then Exit Sub
    
    Dim nSQL As String
    Dim nSQLId As String
    Dim nSQLTmp  As String
    Dim nTitulo As String
    Dim xform As New eps_librerias.FormSeleccion
    Dim xRs As New ADODB.Recordset
    Dim xCampos() As String
    Dim A As Integer
    Dim NUMREGAAGREGAR_ As Integer
    Dim Index As Integer
    Dim IDCRDET_ As Double
    Dim NUMEROMAXTRAB_ As Integer
    Dim IDTAREA_ As Double
    Dim DESCTAREA_ As String
    Dim TAREAACTIVA_ As Boolean
    Dim IDITEM_ As Integer
    Dim DISEÑO_ As Boolean
    
    If CalCtrlCronog.Visible Then
        DISEÑO_ = False
    Else
        DISEÑO_ = True
    End If
    
    Index = 1
    If fg(0).Rows = fg(0).FixedRows Then
        MsgBox "Primero debe procesar tareas, esta operación no esta permitida", _
                vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    If DISEÑO_ Then
        IDCRDET_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_))
        IDITEM_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDITEM_))
    Else
        IDCRDET_ = NulosN(LblIdCrDet.Caption)
        IDITEM_ = NulosN(TxtMatProd.Text)
    End If
            
    If AGREGAR_ Then
        ReDim xCampos(4, 4) As String
        
        xCampos(0, 0) = "Lote":         xCampos(0, 1) = "deslote":      xCampos(0, 2) = "1500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
        xCampos(1, 0) = "Almacen":      xCampos(1, 1) = "desalm":       xCampos(1, 2) = "2500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
        xCampos(2, 0) = "fch. Ing.":    xCampos(2, 1) = "fching":       xCampos(2, 2) = "1000":     xCampos(2, 3) = "D":    xCampos(2, 4) = "D"
        xCampos(3, 0) = "Stock":        xCampos(3, 1) = "cantidad":     xCampos(3, 2) = "1000":     xCampos(3, 3) = "N":    xCampos(3, 4) = "C"
                    
        ' generar la lista de personal para no considerar en la lista
        RstReprocAux.Filter = adFilterNone
        RstReprocAux.Filter = "idcrdet = " & IDCRDET_ & ""
        nSQLId = GENERAR_SQL_ID_RST(RstReprocAux, "idlotedet", " AND alm_inventariolotedet.id", "NOT IN", True)
        
        cSQL = "SELECT alm_inventariolotedet.id AS idlotedet, alm_inventariolote.fching, alm_inventariolote.descripcion AS deslote, alm_inventariolotedet.cantidad, alm_inventariolotedet.idalm, alm_almacenes.descripcion AS desalm " _
            + vbCr + "FROM (alm_inventariolote LEFT JOIN alm_inventariolotedet ON alm_inventariolote.id = alm_inventariolotedet.id) LEFT JOIN alm_almacenes ON alm_inventariolotedet.idalm = alm_almacenes.id " _
            + vbCr + "WHERE (((alm_inventariolotedet.cantidad)>0) AND ((alm_almacenes.tipo)=2) AND ((alm_inventariolote.iditem)=" & IDITEM_ & ")) " & nSQLId _
                    
        nTitulo = "Buscando Reprocesos"
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "fching", "fching", Principio
        If xRs.State = 0 Then Exit Sub
        
        ' agregando los datos al rst temporal
        RstReprocAux.AddNew
        RstReprocAux("idcrdet") = IDCRDET_
        RstReprocAux("idlotedet") = NulosN(xRs("idlotedet"))
        RstReprocAux("deslote") = NulosC(xRs("deslote"))
        RstReprocAux("cantidad") = NulosN(xRs("cantidad"))
        RstReprocAux("idalm") = NulosN(xRs("idalm"))
        RstReprocAux("desalm") = NulosC(xRs("desalm"))
        RstReprocAux.Update
        
        pCargarDatos fg(4), False, False, , , False, DISEÑO_, True
        
        Agregando = False
        Set xform = Nothing
        Set xRs = Nothing
    End If
    
    If LISTAR_ Then
        ReDim xCampos(4, 4) As String
        
        xCampos(0, 0) = "Lote":         xCampos(0, 1) = "deslote":      xCampos(0, 2) = "1500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
        xCampos(1, 0) = "Almacen":      xCampos(1, 1) = "desalm":       xCampos(1, 2) = "2500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
        xCampos(2, 0) = "fch. Ing.":    xCampos(2, 1) = "fching":       xCampos(2, 2) = "1000":     xCampos(2, 3) = "D":    xCampos(2, 4) = "D"
        xCampos(3, 0) = "Stock":        xCampos(3, 1) = "cantidad":     xCampos(3, 2) = "1000":     xCampos(3, 3) = "N":    xCampos(3, 4) = "C"
                  
        ' generar la lista de personal para no considerar en la lista
        RstReprocAux.Filter = "idcrdet = " & IDCRDET_ & ""
        nSQLId = GENERAR_SQL_ID_RST(RstReprocAux, "idlotedet", " AND alm_inventariolotedet.id", "NOT IN", True)
        
        cSQL = "SELECT 0 AS xsel, alm_inventariolotedet.id AS idlotedet, alm_inventariolote.fching, alm_inventariolote.descripcion AS deslote, alm_inventariolotedet.cantidad, alm_inventariolotedet.idalm, alm_almacenes.descripcion AS desalm " _
            + vbCr + "FROM (alm_inventariolote LEFT JOIN alm_inventariolotedet ON alm_inventariolote.id = alm_inventariolotedet.id) LEFT JOIN alm_almacenes ON alm_inventariolotedet.idalm = alm_almacenes.id " _
            + vbCr + "WHERE (((alm_inventariolotedet.cantidad)>0) AND ((alm_almacenes.tipo)=2) AND ((alm_inventariolote.iditem)=" & IDITEM_ & ")) " & nSQLId _
                        
        nTitulo = "Buscando Reprocesos"
    
        xform.SQLCad = cSQL
        xform.titulo = "Buscando Reprocesos"
        Set xform.Coneccion = xCon
        Set xRs = Nothing
        Set xRs = xform.seleccionar(xCampos)
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
                
        xRs.MoveFirst
        While Not xRs.EOF
            ' agregando los datos al rst temporal
            RstReprocAux.AddNew
            RstReprocAux("idcrdet") = IDCRDET_
            RstReprocAux("idlotedet") = NulosN(xRs("idlotedet"))
            RstReprocAux("deslote") = NulosC(xRs("deslote"))
            RstReprocAux("cantidad") = NulosN(xRs("cantidad"))
            RstReprocAux("idalm") = NulosN(xRs("idalm"))
            RstReprocAux("desalm") = NulosC(xRs("desalm"))
            RstReprocAux.Update
            
            xRs.MoveNext
        Wend
        
        RstReprocAux.Filter = adFilterNone
        pCargarDatos fg(4), False, False, , , False, DISEÑO_, True
        
        Set xform = Nothing
        Set xRs = Nothing
    End If
    
    If ELIMINAR_ Then
        If fg(Index).Row < 1 Then
            MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            fg(Index).SetFocus
            Exit Sub
        End If
        
        If fg(Index).Rows = 1 Then
            MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            fg(Index).SetFocus
            Exit Sub
        End If
        
        If Not ELIMINARTODOS_ Then
            If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
        End If
        
        If RstReprocAux.RecordCount <> 0 Then RstReprocAux.MoveFirst
        
        Do While Not RstReprocAux.EOF
            If RstReprocAux.RecordCount = 0 Then Exit Do
            If NulosN(RstReprocAux("idlotedet")) = NulosN(fg(Index).TextMatrix(fg(Index).Row, 4)) Then
                RstReprocAux.Delete
                Exit Do
            End If
            RstReprocAux.MoveNext
        Loop
        
        pCargarDatos fg(Index), False, False, , , False, DISEÑO_, True
    End If
    
    If ELIMTODOS_ Then
        If MsgBox("¿Esta seguro de eliminar todos los registros?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
        ELIMINARTODOS_ = True
        
        For A = 1 To fg(Index).Rows - 1
            Agregando = False
            fg(Index).Select 1, 1, 1, fg(Index).Cols - 1
            procesarReproceso False, False, True, False
        Next A
        ELIMINARTODOS_ = False
        pCargarDatos fg(Index), False, False, , , False, DISEÑO_, True
    End If
End Sub

Private Sub procesarPersonal(AGREGAR_ As Boolean, LISTAR_ As Boolean, _
                                    ELIMINAR_ As Boolean, ELIMTODOS_ As Boolean, _
                                    Optional DISEÑO_ As Boolean = False)
    If QueHace = 3 Then Exit Sub
    
    Dim nSQL As String
    Dim nSQLId As String
    Dim nSQLTmp  As String
    Dim nTitulo As String
    Dim xform As New eps_librerias.FormSeleccion
    Dim xRs As New ADODB.Recordset
    Dim xCampos() As String
    Dim A As Integer
    Dim NUMREGAAGREGAR_ As Integer ' numero de registros que se van a agregar
    
    Dim Index As Integer
    Dim IDCRDET_ As Double
    Dim NUMEROMAXTRAB_ As Integer
    Dim IDTAREA_ As Double
    Dim DESCTAREA_ As String
    Dim TAREAACTIVA_ As Boolean
    Dim IDAREA_ As Double
    
    Index = 1
    If fg(0).Rows = fg(0).FixedRows Then
        MsgBox "Primero debe procesar tareas, esta operación no esta permitida", _
                vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    If DISEÑO_ Then
        IDCRDET_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_))
        NUMEROMAXTRAB_ = NulosN(fg(0).TextMatrix(fg(0).Row, 6))
        IDTAREA_ = NulosN(fg(0).TextMatrix(fg(0).Row, 11))
        DESCTAREA_ = NulosC(fg(0).TextMatrix(fg(0).Row, 2))
        TAREAACTIVA_ = fg(0).TextMatrix(fg(0).Row, 1)
        IDAREA_ = fg(0).TextMatrix(fg(0).Row, 16)
    Else
        IDCRDET_ = NulosN(LblIdCrDet.Caption)
        NUMEROMAXTRAB_ = NulosN(lblntrab.Caption)
        IDTAREA_ = NulosN(fg(0).TextMatrix(fg(0).Row, 11))
        DESCTAREA_ = NulosC(fg(0).TextMatrix(fg(0).Row, 2))
        TAREAACTIVA_ = fg(0).TextMatrix(fg(0).Row, 1)
        IDAREA_ = fg(0).TextMatrix(fg(0).Row, 16)
    End If
    
    If Not TAREAACTIVA_ Then
        MsgBox "La Tarea actual no esta activa, esta operación no esta permitida", _
                vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
        
    If AGREGAR_ Then
        ReDim xCampos(5, 4) As String
        
        xCampos(0, 0) = "DNI":                  xCampos(0, 1) = "numdoc":      xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
        xCampos(1, 0) = "Grupo":                xCampos(1, 1) = "grupo":       xCampos(1, 2) = "800":      xCampos(1, 3) = "N":    xCampos(1, 4) = "C"
        xCampos(2, 0) = "Apellidos y Nombres":  xCampos(2, 1) = "nombre":      xCampos(2, 2) = "3250":     xCampos(2, 3) = "C":    xCampos(2, 4) = "C"
        xCampos(3, 0) = "Area":                 xCampos(3, 1) = "area":        xCampos(3, 2) = "1750":     xCampos(3, 3) = "C":    xCampos(3, 4) = "C"
        xCampos(4, 0) = "Fch. Ing.":            xCampos(4, 1) = "fching":      xCampos(4, 2) = "1000":     xCampos(4, 3) = "C":    xCampos(4, 4) = "C"
        
        
        If LIMITARNUMEROPERSONAL_ Then
            If fg(Index).Rows - 1 >= NUMEROMAXTRAB_ Then
                MsgBox "La Tarea actual no puede admitir mas Personal", _
                        vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
        End If
            
        ' generar la lista de personal para no considerar en la lista
        RstPersonalAux.Filter = "idcrdet = " & IDCRDET_ & ""
        nSQLId = GENERAR_SQL_ID_RST(RstPersonalAux, "idper", " AND pla_empleados.id", "NOT IN", True)
        
        If LIMITARSELPERSONAL_ Then
            ' generar la consulta
            nSQL = "SELECT pla_empleados.numdoc, pro_grupo.num AS grupo, pla_empleados.id AS idemp, pla_empleados.nombre, -1 AS activo, mae_area.descripcion AS area, pla_empleados.fching " _
                + vbCr + "FROM (((pla_empleados LEFT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN mae_area ON pla_empleados.idarea = mae_area.id) LEFT JOIN (pro_grupodet LEFT JOIN pro_grupo ON pro_grupodet.idgrupo = pro_grupo.id) ON pro_emp.id = pro_grupodet.idper " _
                + vbCr + "Where (((pla_empleados.fchcese) Is Null) And ((pro_empdet.idfun) = 6) And ((pla_empleados.idarea) = " & IDAREA_ & ")) " & nSQLId _
                + vbCr + "ORDER BY pla_empleados.nombre;"
        Else
            ' generar la consulta
            nSQL = "SELECT pla_empleados.numdoc, pro_grupo.num AS grupo, pla_empleados.id AS idemp, pla_empleados.nombre, -1 AS activo, mae_area.descripcion AS area, pla_empleados.fching " _
                + vbCr + "FROM (((pla_empleados LEFT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN mae_area ON pla_empleados.idarea = mae_area.id) LEFT JOIN (pro_grupodet LEFT JOIN pro_grupo ON pro_grupodet.idgrupo = pro_grupo.id) ON pro_emp.id = pro_grupodet.idper " _
                + vbCr + "Where (((pla_empleados.fchcese) Is Null) And ((pro_empdet.idfun) = 6)) " & nSQLId _
                + vbCr + "ORDER BY pla_empleados.nombre;"
        End If
                    
        nTitulo = "Buscando Personal"
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
            
        xform.titulo = "Buscando Personal"
        
        If xRs.State = 0 Then Exit Sub
        
        If fg(Index).Rows = fg(Index).FixedRows Then fg(Index).Rows = fg(Index).Rows + 1
        
        ' agregando los datos al rst temporal
        RstPersonalAux.AddNew
        RstPersonalAux("idcrdet") = IDCRDET_
        RstPersonalAux("idtar") = IDTAREA_
        RstPersonalAux("destar") = DESCTAREA_
        RstPersonalAux("activo") = NulosN(xRs("activo"))
        RstPersonalAux("idper") = NulosN(xRs("idemp"))
        RstPersonalAux("nombre") = NulosC(xRs("nombre"))
        RstPersonalAux("numdoc") = NulosC(xRs("numdoc"))
        RstPersonalAux.Update
        
        pCargarDatos fg(Index), True, False, , , False, DISEÑO_
        
        Agregando = False
        Set xform = Nothing
        Set xRs = Nothing
    End If
    
    If LISTAR_ Then
        If LIMITARNUMEROPERSONAL_ Then
            NUMREGAAGREGAR_ = NUMEROMAXTRAB_ - (fg(Index).Rows - 1)
            If NUMREGAAGREGAR_ <= 0 Then
                MsgBox "La Tarea actual no puede admitir mas Personal", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
        End If
        
        ReDim xCampos(5, 4) As String
        
        xCampos(0, 0) = "DNI":                  xCampos(0, 1) = "numdoc":      xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
        xCampos(1, 0) = "Grupo":                xCampos(1, 1) = "grupo":       xCampos(1, 2) = "800":      xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
        xCampos(2, 0) = "Apellidos y Nombres":  xCampos(2, 1) = "nombre":      xCampos(2, 2) = "3250":     xCampos(2, 3) = "C":    xCampos(2, 4) = "C"
        xCampos(3, 0) = "Area":                 xCampos(3, 1) = "area":        xCampos(3, 2) = "1750":     xCampos(3, 3) = "C":    xCampos(3, 4) = "C"
        xCampos(4, 0) = "Fch. Ing.":            xCampos(4, 1) = "fching":      xCampos(4, 2) = "1000":     xCampos(4, 3) = "D":    xCampos(4, 4) = "C"
                  
        ' generar la lista de personal para no considerar en la lista
        RstPersonalAux.Filter = "idcrdet = " & IDCRDET_ & ""
        nSQLId = GENERAR_SQL_ID_RST(RstPersonalAux, "idper", " AND pla_empleados.id", "NOT IN", True)
        
        If LIMITARSELPERSONAL_ Then
            ' generar la consulta
            nSQL = "SELECT 0 AS xsel, pla_empleados.numdoc, pro_grupo.num AS grupo, pla_empleados.id AS idemp, pla_empleados.nombre, -1 AS activo, mae_area.descripcion AS area, pla_empleados.fching " _
                + vbCr + "FROM (((pla_empleados LEFT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN mae_area ON pla_empleados.idarea = mae_area.id) LEFT JOIN (pro_grupodet LEFT JOIN pro_grupo ON pro_grupodet.idgrupo = pro_grupo.id) ON pro_emp.id = pro_grupodet.idper " _
                + vbCr + "Where (((pla_empleados.fchcese) Is Null) And ((pro_empdet.idfun) = 6) And ((pla_empleados.idarea) = " & IDAREA_ & ")) " & nSQLId _
                + vbCr + "ORDER BY pla_empleados.nombre;"
        Else
            ' generar la consulta
            nSQL = "SELECT 0 AS xsel, pla_empleados.numdoc, pro_grupo.num AS grupo, pla_empleados.id AS idemp, pla_empleados.nombre, -1 AS activo, mae_area.descripcion AS area, pla_empleados.fching " _
                + vbCr + "FROM (((pla_empleados LEFT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN mae_area ON pla_empleados.idarea = mae_area.id) LEFT JOIN (pro_grupodet LEFT JOIN pro_grupo ON pro_grupodet.idgrupo = pro_grupo.id) ON pro_emp.id = pro_grupodet.idper " _
                + vbCr + "Where (((pla_empleados.fchcese) Is Null) And ((pro_empdet.idfun) = 6)) " & nSQLId _
                + vbCr + "ORDER BY pla_empleados.nombre;"
        End If
            
        nTitulo = "Buscando Personal"
    
        xform.SQLCad = nSQL
            
        xform.titulo = "Buscando Personal"
        Set xform.Coneccion = xCon
        Set xRs = Nothing
        Set xRs = xform.seleccionar(xCampos)
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
          
        If fg(Index).Rows = fg(Index).FixedRows Then fg(Index).Rows = fg(Index).Rows + 1
        
        If Not LIMITARNUMEROPERSONAL_ Then NUMREGAAGREGAR_ = xRs.RecordCount
        For A = 1 To NUMREGAAGREGAR_
            ' agregando los datos al rst temporal
            RstPersonalAux.AddNew
            RstPersonalAux("idcrdet") = IDCRDET_
            RstPersonalAux("idtar") = IDTAREA_
            RstPersonalAux("destar") = DESCTAREA_
            RstPersonalAux("activo") = xRs("activo")
            RstPersonalAux("idper") = xRs("idemp")
            RstPersonalAux("nombre") = xRs("nombre")
            RstPersonalAux("numdoc") = xRs("numdoc")
            
            xRs.MoveNext
            If xRs.EOF = True Then Exit For
        Next A
        
        RstPersonalAux.Filter = adFilterNone
        pCargarDatos fg(Index), True, False, , , False, DISEÑO_
        
        Set xform = Nothing
        Set xRs = Nothing
    End If
    
    If ELIMINAR_ Then
        If fg(Index).Row < 1 Then
            MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            fg(Index).SetFocus
            Exit Sub
        End If
        
        If fg(Index).Rows = 1 Then
            MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            fg(Index).SetFocus
            Exit Sub
        End If
        
        If Not ELIMINARTODOS_ Then
            If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
        End If
        
        If RstPersonalAux.RecordCount <> 0 Then RstPersonalAux.MoveFirst
        
        Do While Not RstPersonalAux.EOF
            If RstPersonalAux.RecordCount = 0 Then Exit Do
            If NulosN(RstPersonalAux("idper")) = NulosN(fg(Index).TextMatrix(fg(Index).Row, 5)) Then
                RstPersonalAux.Delete
                Exit Do
            End If
            RstPersonalAux.MoveNext
        Loop
        
        pCargarDatos fg(Index), True, False, , , False, DISEÑO_
    End If
    
    If ELIMTODOS_ Then
        If MsgBox("¿Esta seguro de eliminar todos los registros?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
        ELIMINARTODOS_ = True
        
        For A = 1 To fg(Index).Rows - 1
            Agregando = False
            fg(Index).Select 1, 1, 1, fg(Index).Cols - 1
            procesarPersonal False, False, True, False, DISEÑO_
        Next A
        ELIMINARTODOS_ = False
        pCargarDatos fg(Index), True, False, , , False, DISEÑO_
    End If
    
    totalizarPersonal
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
        LIMITARNUMEROPERSONAL_ = cknumper.Value
        LIMITARNUMEROTAREAS_ = cknumtar.Value
        LIMITARSELPERSONAL_ = ckperarea.Value
    End If
    
    If CARGAR_ Then
        OptTarea(MODO_TAREA).Value = True
        OptHoras(INCLUIR_HORAS).Value = True
        TxtPctje.Text = PORCENTAJE
        DTPMinutos.Value = MINUTOS_
        DTPHorIni.Value = HOR_INI
        DTPHorFin.Value = HOR_FIN
        
        If LIMITARNUMEROPERSONAL_ Then cknumper.Value = 1 Else cknumper.Value = 0
        If LIMITARNUMEROTAREAS_ Then cknumtar.Value = 1 Else cknumtar.Value = 0
        If LIMITARSELPERSONAL_ Then ckperarea.Value = 1 Else ckperarea.Value = 0
    End If
End Sub

Private Sub agregarCampos(PRODUCTO_ As Boolean, TAREA_ As Boolean, _
                        Optional RESPONSABLE_ As Boolean = False, _
                        Optional RECETA_ As Boolean = False, _
                        Optional LINEA_ As Boolean = False, _
                        Optional DISEÑO_ As Boolean = False)
    Dim xCampos() As String
    Dim RstLinea As New ADODB.Recordset
    Dim nTitulo As String
    Dim RstTmp As New ADODB.Recordset
    Dim IDCRDET_ As Double
    
    If PRODUCTO_ Then
        ReDim xCampos(2, 4) As String
        Dim xRs As New ADODB.Recordset
        Dim titulo As String
        Dim Rpta As Integer
        
        If QueHace = 3 Then Exit Sub
    
        'descripcion                     'campo                       'tamaño                         'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "despro":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Uni. Med.":     xCampos(1, 1) = "abrev":     xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"

        cSQL = "SELECT pro_receta.iditem, pro_receta.id AS idrec, pro_receta.codrec, alm_inventario.descripcion AS despro, mae_unidades.abrev, pro_tiptrab.id AS idtiptrab, pro_tiptrab.descripcion AS destiptrab, pro_formapag.id AS idformapag, pro_formapag.descripcion AS desformapag " _
            + vbCr + "FROM (((pro_receta LEFT JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_receta.idunimed = mae_unidades.id) LEFT JOIN pro_tiptrab ON pro_receta.idtiptrab = pro_tiptrab.id) LEFT JOIN pro_formapag ON pro_receta.idformapag = pro_formapag.id " _
            + vbCr + "WHERE (((pro_receta.prirec)=1) AND ((alm_inventario.activo)=-1));"
            
        titulo = "Buscando Productos"
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos, titulo, "despro", "despro"
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        If DISEÑO_ Then
            IDCRDET_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_))
        Else
            IDCRDET_ = NulosN(LblIdCrDet.Caption)
        End If
            
        RstTareasAux.Filter = "idcrdet = " & IDCRDET_
        
        If RstTareasAux.RecordCount > 0 Then
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
        If DISEÑO_ Then
            fg(3).TextMatrix(fg(3).Row, COLUMNAPRODUCTO_) = NulosC(xRs("despro"))       ' Descripcion
            fg(3).TextMatrix(fg(3).Row, COLUMNARECETA_) = NulosC(xRs("codrec"))         ' Receta
            fg(3).TextMatrix(fg(3).Row, COLUMNAUM_) = NulosC(xRs("abrev"))              ' UM
            fg(3).TextMatrix(fg(3).Row, COLUMNAIDRECETA_) = NulosN(xRs("idrec"))        ' Idreceta
            fg(3).TextMatrix(fg(3).Row, COLUMNAIDITEM_) = NulosN(xRs("iditem"))         ' Iditem
        Else
            ' Producto
            TxtMatProd.Text = NulosN(xRs("iditem"))
            LblMatProd.Caption = NulosC(xRs("despro"))
            ' Unidad
            LblUnidad.Caption = NulosC(xRs("abrev"))
            ' Receta
            TxtCodRec.Text = NulosC(xRs("codrec"))
            LblIdRec.Caption = NulosN(xRs("idrec"))
            
            Cmd(18).SetFocus
        End If

        ' Se verifica si el producto seleccionado tiene una linea activa
        cSQL = "SELECT pro_linea.id AS idlineadet, pro_linea.descripcion " _
                + vbCr + "From pro_linea " _
                + vbCr + "WHERE (((pro_linea.idrec)=" & NulosN(xRs("idrec")) & ") AND ((pro_linea.activo)=-1));"
                        
        RST_Busq RstLinea, cSQL, xCon
        
        If RstLinea.State = 0 Then GoTo ERROR_AL_ENCONTRAR_LINEA
        If RstLinea.RecordCount = 0 Then GoTo ERROR_AL_ENCONTRAR_LINEA
        
        ' Se llena la linea Activa
        If DISEÑO_ Then
            fg(3).TextMatrix(fg(3).Row, COLUMNALINEA_) = NulosC(RstLinea("descripcion"))  ' Linea
            fg(3).TextMatrix(fg(3).Row, COLUMNAIDLINEA_) = NulosN(RstLinea("idlineadet"))  ' Idlinea
        Else
            TxtIdLineaDet.Text = NulosN(RstLinea("idlineadet"))
            LblLinea.Caption = NulosC(RstLinea("descripcion"))
        End If
        Set xRs = Nothing
        Set RstLinea = Nothing
        Exit Sub
        
ERROR_AL_ENCONTRAR_LINEA:
        MsgBox "El producto procesado no tiene Linea activa, procese una para calcular tiempos de Producción", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        If DISEÑO_ Then
            fg(3).TextMatrix(fg(3).Row, COLUMNALINEA_) = ""
            fg(3).TextMatrix(fg(3).Row, COLUMNAIDLINEA_) = 0
        Else
            TxtIdLineaDet.Text = ""
            LblLinea.Caption = ""
        End If
        
        Set xRs = Nothing
        Set RstLinea = Nothing
    End If
    
    If TAREA_ Then
    
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
        
        If DISEÑO_ Then
            fg(3).TextMatrix(fg(3).Row, COLUMNAENCARGADO_) = NulosC(xRs("nombre"))  ' Responsable
            fg(3).TextMatrix(fg(3).Row, COLUMNAIDRESP_) = NulosN(xRs("idemp"))  ' idresponsable
        Else
            LblEncargado.Caption = NulosC(xRs("nombre"))     ' Responsable
            TxtIdEncarg.Text = NulosN(xRs("idemp"))          ' idresponsable
            TxtCant.SetFocus
        End If
    End If
    
    If RECETA_ Then ' Cargar Receta
        ReDim xCampos(2, 4) As String
        Dim IDITEM_ As Double
        
        If DISEÑO_ Then
            IDITEM_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDITEM_))
        Else
            IDITEM_ = NulosN(TxtMatProd.Text)
        End If
        
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Receta":     xCampos(1, 1) = "codrec":        xCampos(1, 2) = "1200":         xCampos(1, 3) = "C"
        
        cSQL = "SELECT pro_receta.codrec, pro_receta.descripcion, pro_receta.prirec, pro_receta.id " _
            + vbCr + "From pro_receta " _
            + vbCr + "Where (((pro_receta.iditem) = " & IDITEM_ & ")) " _
            + vbCr + "ORDER BY pro_receta.prirec;"
            
        nTitulo = "Buscando Recetas del Producto"
                
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        If DISEÑO_ Then
            fg(3).TextMatrix(fg(3).Row, COLUMNARECETA_) = NulosC(xRs("codrec"))  ' Receta
            fg(3).TextMatrix(fg(3).Row, COLUMNAIDRECETA_) = NulosN(xRs("id")) ' idreceta
        Else
            TxtCodRec.Text = NulosC(xRs("codrec"))             ' Codigo de la receta
            LblIdRec.Caption = NulosN(xRs("id"))               ' Id de la receta
        End If
        
        ' Se verifica si el producto seleccionado tiene una linea activa
        cSQL = "SELECT pro_linea.id AS idlineadet, pro_linea.descripcion " _
                + vbCr + "From pro_linea " _
                + vbCr + "WHERE (((pro_linea.idrec)=" & NulosN(xRs("id")) & ") AND ((pro_linea.activo)=-1));"
                        
        RST_Busq RstLinea, cSQL, xCon
        
        If RstLinea.State = 0 Then GoTo ERROR_AL_ENCONTRAR_LINEA2
        If RstLinea.RecordCount = 0 Then GoTo ERROR_AL_ENCONTRAR_LINEA2
        
        ' Se llena la linea Activa
        If DISEÑO_ Then
            fg(3).TextMatrix(fg(3).Row, COLUMNALINEA_) = NulosC(RstLinea("descripcion"))  ' Linea
            fg(3).TextMatrix(fg(3).Row, COLUMNAIDLINEA_) = NulosN(RstLinea("idlineadet"))  ' idlinea
        Else
            TxtIdLineaDet.Text = NulosN(RstLinea("idlineadet"))
            LblLinea.Caption = NulosC(RstLinea("descripcion"))
        End If
        
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
        Dim IDREC_ As Double
        
        If DISEÑO_ Then
            IDREC_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDRECETA_))
        Else
            IDREC_ = NulosN(LblIdRec.Caption)
        End If
        
        'descripcion                            'campo                          'tamaño                        'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripcion":          xCampos(0, 1) = "descline":     xCampos(0, 2) = "4000":        xCampos(0, 3) = "C"
        xCampos(1, 0) = "Operarios":            xCampos(1, 1) = "numop":        xCampos(1, 2) = "1000":        xCampos(1, 3) = "N"
        xCampos(2, 0) = "Eficiencia (%)":       xCampos(2, 1) = "efic":         xCampos(2, 2) = "1250":        xCampos(2, 3) = "N"
     
        cSQL = "SELECT pro_linea.descripcion AS descline, pro_linea.numop, pro_linea.efic, pro_linea.idlinea, pro_linea.id AS idlineadet " _
            + vbCr + "From pro_linea " _
            + vbCr + "WHERE (((pro_linea.idrec)=" & IDREC_ & "));"
    
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
        If DISEÑO_ Then
            fg(3).TextMatrix(fg(3).Row, COLUMNALINEA_) = NulosC(RstTmp("descline"))  ' Linea
            fg(3).TextMatrix(fg(3).Row, COLUMNAIDLINEA_) = NulosN(RstTmp("idlineadet"))  ' idlinea
        Else
            TxtIdLineaDet.Text = NulosN(RstTmp("idlineadet"))
            LblLinea.Caption = NulosC(RstTmp("descline"))
            Cmd(2).SetFocus
        End If
        
        Set RstTmp = Nothing
    End If
End Sub

Private Sub pCargarDatos(Fgrid As VSFlexGrid, _
                        Optional PERSONAL_ As Boolean = True, _
                        Optional TAREAS_ As Boolean = False, _
                        Optional TODOS_ As Boolean = False, _
                        Optional RECETA_ As Boolean = False, _
                        Optional NUEVO_ As Boolean = True, _
                        Optional DISEÑO_ As Boolean = False, _
                        Optional REPROCESO_ As Boolean = False)
    
    Dim A As Integer
    Dim IDCRDET_ As Double
    Dim IDTAR_ As Double
    
    If DISEÑO_ Then
        IDCRDET_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_))
'        If fg(5).Rows > fg(5).FixedRows Then IDTAR_ = NulosN(fg(5).TextMatrix(fg(5).Row, 11)) _
'        Else IDTAR_ = 0
    Else
        IDCRDET_ = NulosN(LblIdCrDet.Caption)
'        If fg(0).Rows > fg(0).FixedRows Then IDTAR_ = NulosN(fg(0).TextMatrix(fg(0).Row, 11)) _
'        Else IDTAR_ = 0
    End If
    
    '******************************************************************************************
    If fg(0).Rows > fg(0).FixedRows Then IDTAR_ = NulosN(fg(0).TextMatrix(fg(0).Row, 11)) _
    Else IDTAR_ = 0
    '******************************************************************************************
    
    Agregando = True
        
    With Fgrid
        If PERSONAL_ Then ' Si se desea cargar personal
            .Rows = 1
            If RstPersonal.State = 0 Then Agregando = False: Exit Sub
            
            If NUEVO_ Then
                RstPersonal.Filter = adFilterNone
                RstPersonal.Filter = "idcrdet = " & IDCRDET_ & ""
                ' Se verifica que este creado el recordset
                If RstPersonalAux.State = 0 Then DEFINIR_RST_TMP RstPersonalAux, RstPersonal
                ' Se vacia el recordset
                limpiarRST RstPersonalAux
                ' Se carga con los datos temporales
                CARGAR_RST_TMP RstPersonalAux, RstPersonal
            End If
            
            If TODOS_ Then ' si se muestran todos los trabajadores
                RstPersonalAux.Filter = adFilterNone
                RstPersonalAux.Filter = "idcrdet = " & IDCRDET_ & ""
            Else ' si se muestran solo de una tarea especifica
                RstPersonalAux.Filter = "idcrdet = " & IDCRDET_ & " And idtar = " & IDTAR_ & ""
            End If
            
            If RstPersonalAux.RecordCount = 0 Then ' Si no hay Personal
                '*******************************************************************
                LblDetTrab.Caption = (.Rows - 1) & " de " & NulosN(lblntrab.Caption)
                Agregando = False
                '*******************************************************************
                Exit Sub
            End If
            
            ' Se llena al Personal
            Do While Not RstPersonalAux.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = NulosN(RstPersonalAux.Fields("activo"))
                .TextMatrix(.Rows - 1, 2) = NulosC(RstPersonalAux.Fields("numdoc"))
                .TextMatrix(.Rows - 1, 3) = NulosC(RstPersonalAux.Fields("nombre"))
                .TextMatrix(.Rows - 1, 4) = NulosN(RstPersonalAux.Fields("idcrdet"))
                .TextMatrix(.Rows - 1, 5) = NulosN(RstPersonalAux.Fields("idper"))
                .TextMatrix(.Rows - 1, 6) = NulosN(RstPersonalAux.Fields("idtar"))
                .TextMatrix(.Rows - 1, 7) = NulosC(RstPersonalAux.Fields("destar"))
                
                RstPersonalAux.MoveNext
            Loop
            
            '*****************************************************************
            LblDetTrab.Caption = .Rows - 1 & " de " & NulosN(lblntrab.Caption)
            '*****************************************************************
            
            ' aplicando el orden a la lista de datos
            GRID_ORDENAR Fgrid, 1, 2
        End If
        
        If TAREAS_ Then ' Si se desea cargar Tareas
            .Rows = 1
            ' Si no hay Tareas guardadas
            If RstTareas.State = 0 Then Agregando = False: Exit Sub
            ' Se verfica si es una carga nueva o actualizacion de datos
            If NUEVO_ Then
                ' Se filtra el registro involucrado
                RstTareas.Filter = "idcrdet = " & IDCRDET_ & ""
                If RstTareasAux.State = 0 Then DEFINIR_RST_TMP RstTareasAux, RstTareas
                limpiarRST RstTareasAux
                CARGAR_RST_TMP RstTareasAux, RstTareas
            End If
            
            RstTareasAux.Filter = "idcrdet = " & IDCRDET_ & ""
            If RstTareasAux.RecordCount = 0 Then Agregando = False: Exit Sub
            
            
            Dim PRIMERAFILAACTIVA_ As Integer
            
            PRIMERAFILAACTIVA_ = 0
            ' Se procede a llenar las tareas
            RstTareasAux.MoveFirst
            While Not RstTareasAux.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = NulosN(RstTareasAux.Fields("activo"))
                
                ' Se busca la primera fila activa
                If PRIMERAFILAACTIVA_ = 0 Then
                    If NulosN(RstTareasAux.Fields("activo")) = -1 Then
                        PRIMERAFILAACTIVA_ = .Rows - 1
                    End If
                End If
                
                .TextMatrix(.Rows - 1, 2) = NulosC(RstTareasAux.Fields("destar"))
                .TextMatrix(.Rows - 1, 3) = Format(RstTareasAux.Fields("durtar"), "HH:mm")
                .TextMatrix(.Rows - 1, 4) = Format(RstTareasAux.Fields("horinitar"), "HH:mm")
                .TextMatrix(.Rows - 1, 5) = Format(RstTareasAux.Fields("horfintar"), "HH:mm")
                .TextMatrix(.Rows - 1, 6) = Format(NulosN(RstTareasAux.Fields("numper")), "00")
                .TextMatrix(.Rows - 1, 7) = Format(NulosN(RstTareasAux.Fields("cantproc")), "0.00")
                .TextMatrix(.Rows - 1, 8) = Format(NulosC(RstTareasAux.Fields("fchini")), FORMAT_DATE)
                .TextMatrix(.Rows - 1, 9) = Format(NulosC(RstTareasAux.Fields("fchfin")), FORMAT_DATE)
                .TextMatrix(.Rows - 1, 10) = NulosN(RstTareasAux.Fields("idcrdet"))
                .TextMatrix(.Rows - 1, 11) = NulosN(RstTareasAux.Fields("idtar"))
                .TextMatrix(.Rows - 1, 12) = NulosN(RstTareasAux.Fields("aplpor"))
                .TextMatrix(.Rows - 1, 13) = NulosC(RstTareasAux.Fields("desarea"))
                .TextMatrix(.Rows - 1, 14) = NulosC(RstTareasAux.Fields("destiptrab"))
                .TextMatrix(.Rows - 1, 16) = NulosN(RstTareasAux.Fields("idarea"))
                .TextMatrix(.Rows - 1, 17) = NulosN(RstTareasAux.Fields("idtiptrab"))
                
                '**********************************************************************
                .TextMatrix(.Rows - 1, 15) = NulosC(RstTareasAux.Fields("nomresp"))
                .TextMatrix(.Rows - 1, 18) = NulosN(RstTareasAux.Fields("idresp"))
                '**********************************************************************
                
                If NulosN(RstTareasAux.Fields("activo")) = True Then
                    .Select .Rows - 1, 1, .Rows - 1, .Cols - 1
                    .FillStyle = flexFillRepeat
                    .CellBackColor = &HFFFF&
                End If
                
                RstTareasAux.MoveNext
            Wend
            .TopRow = PRIMERAFILAACTIVA_
            .Select PRIMERAFILAACTIVA_, 2
        End If
        
        If RECETA_ Then ' No disponible
        End If
        
        If REPROCESO_ Then
            .Rows = 1
            ' Si no hay Tareas guardadas
            If RstReproc.State = 0 Then Agregando = False: Exit Sub
            ' Se verfica si es una carga nueva o actualizacion de datos
            If NUEVO_ Then
                ' Se filtra el registro involucrado
                RstReproc.Filter = "idcrdet = " & IDCRDET_ & ""
                If RstReprocAux.State = 0 Then DEFINIR_RST_TMP RstReprocAux, RstReproc
                limpiarRST RstReprocAux
                CARGAR_RST_TMP RstReprocAux, RstReproc
            End If
            
            RstReprocAux.Filter = "idcrdet = " & IDCRDET_ & ""
            If RstReprocAux.RecordCount = 0 Then Agregando = False: Exit Sub
                       
            ' Se procede a llenar las tareas
            RstReprocAux.MoveFirst
            While Not RstReprocAux.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = NulosN(RstReprocAux("deslote"))
                .TextMatrix(.Rows - 1, 2) = NulosC(RstReprocAux("desalm"))
                .TextMatrix(.Rows - 1, 3) = Format(NulosN(RstReprocAux("cantidad")), FORMAT_CANTIDAD)
                .TextMatrix(.Rows - 1, 4) = NulosN(RstReprocAux.Fields("idlotedet"))
                .TextMatrix(.Rows - 1, 5) = NulosN(RstReprocAux.Fields("idalm"))
                RstReprocAux.MoveNext
            Wend
            
            ' aplicando el orden a la lista de datos
            GRID_ORDENAR Fgrid, 1, 1
        End If
    End With
    
    Agregando = False
End Sub

Private Sub calcularDatosAdicionales(Optional DISEÑO_ As Boolean)
    Dim A As Integer
    Dim HORAFIN_ As Date
    Dim FCHFIN_ As String
    Dim CANPRO_ As Double
    Dim IDCRDET_ As Double
    Dim NUMOPE_ As Double
    Dim NUMOPESEL_ As Double
    
    If DISEÑO_ Then
        IDCRDET_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_))
    Else
        IDCRDET_ = NulosN(LblIdCrDet.Caption)
    End If
    
    ' Se filtran las Tareas activas
    If RstTareasAux.State = 0 Then Exit Sub
    RstTareasAux.Filter = "idcrdet = " & IDCRDET_ & " And activo = -1"
    
    If RstTareasAux.RecordCount = 0 Then
        HORAFIN_ = 0
        CANPRO_ = 0
        GoTo SALIR
    End If
    
    ' Se Busca la primera Tarea Seleccionada
    ' para llenar la cantidad de Mp
    RstTareasAux.MoveFirst
    NUMOPE_ = 0
    If NulosN(RstTareasAux("aplpor")) = 0 Then RstTareasAux("aplpor") = 100
    CANPRO_ = (NulosN(RstTareasAux("cantproc")) * 100) / NulosN(RstTareasAux("aplpor"))
    
    ' Numero de Operarios
    While Not RstTareasAux.EOF
        NUMOPE_ = NUMOPE_ + NulosN(RstTareasAux("numper"))
        RstTareasAux.MoveNext
    Wend
    ' Se Busca la Ultima Tarea Seleccionada
    ' para llenar la fecha y hora de fin
    RstTareasAux.MoveLast
    HORAFIN_ = RstTareasAux("horfintar")
    FCHFIN_ = RstTareasAux("fchfin")

SALIR:
    RstPersonalAux.Filter = "idcrdet = " & IDCRDET_
    NUMOPESEL_ = RstPersonalAux.RecordCount
    
    lblFchFin.Caption = Format(FCHFIN_, FORMAT_DATE)
    LblHorFin.Caption = Format(HORAFIN_, FORMAT_HORA_SIN_SEGUNDO)
    lblntrabtot.Caption = NUMOPE_
    lblNumOpe.Caption = Format(NUMOPESEL_, "00") & " de " & Format(NUMOPE_, "00")
    
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
                        Optional NOMRESPONSABLE_ As String = "", _
                        Optional PORCENTAJEAPLICADO_ As Double = 100)

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

    If RstTareas_Aux.State = 0 Then Exit Sub
    If RstTareas_Aux.RecordCount = 0 Then Exit Sub
    
    RstTareas_Aux.MoveFirst
    
    Agregando = True
    
    Dim xRs As New ADODB.Recordset
    
    DEFINIR_RST_TMP xRs, RstTareas_Aux
    CARGAR_RST_TMP xRs, RstTareas_Aux
    
    RstTareas_Aux.MoveFirst
    
    ' Se halla los valores iniciales de los campos cuando no es un ingreso nuevo
    cantidad_procesada_anterior = cantidad_procesada
    hora_inicio_tarea_anterior = hora_inicio
    hora_fin_tarea_anterior = hora_fin
    duracion_tarea_anterior = Format(CDate(hora_fin) - CDate(hora_inicio), "HH:mm")
    
    fecha_fin_tarea = fecha_fin
    fecha_Inicio_Tarea = fecha_fin
    
    ' Se proceden a procesar y agregar todos los productos filtrados
    For B = 1 To RstTareas_Aux.RecordCount
        If RstTareas_Aux("activo") = 0 Then GoTo SIGUIENTE
        
        If es_nuevo Then ' Si es nuevo se agrega un nuevo registro al Recordset de Tareas
            RstTareasAux.AddNew
        Else ' Sino se filtra el registro involucrado
            RstTareasAux.Filter = "idcrdet = " & ID_CRDET_ & " And idtar = " & RstTareas_Aux("idtar") & ""
        End If

        ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>Se llena la cantidad porcentual
        ' Si el procentaje aplicado a la tarea es cero
        Dim PORCENTAJE_AUX As Double
        
        PORCENTAJE_AUX = NulosN(RstTareas_Aux("aplpor"))
        If PORCENTAJE_AUX = 0 Then PORCENTAJE_AUX = 100
        CANTIDAD_ = cantidad_procesada_anterior
        
        RstTareasAux("cantproc") = (NulosN(CANTIDAD_) * ((PORCENTAJE_AUX / 100)))
        
        ' Se actualiza la cantidad que se va a procesar
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
        
        ' Se aplica el cambio porcentual de la duracion de la Tarea
        If PORCENTAJEAPLICADO_ = 0 Then PORCENTAJEAPLICADO_ = 100
        FACTOR_ = (FACTOR_ * 100) / PORCENTAJEAPLICADO_
        
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
        
        Select Case Tipo
            Case 0 ' una tarea despues de otra
                RstTareasAux("horinitar") = hora_fin_tarea_anterior
    
            Case 1 ' una tarea al porcentaje de otra
                ' Se aplica el porcentaje
                h = Split(duracion_tarea_anterior, ":")
                tiempo = (60 * Val(h(0))) + Val(h(1))
                tiempo = ((valor * tiempo) / 100)
                tiempo = tiempo / 60 ' Se cambia a horas
    
                intervalo = Format(Int(tiempo), "00")
                intervalo = intervalo & ":" & Format(((tiempo * 60) Mod 60), "00")
                RstTareasAux("horinitar") = CDate(hora_inicio_tarea_anterior) + CDate(intervalo)

            Case 2 ' Una tarea al minuto de otra
                If hora_inicio_tarea_anterior = hora_fin_tarea_anterior Then
                    RstTareasAux("horinitar") = hora_inicio_tarea_anterior
                Else
                    RstTareasAux("horinitar") = CDate(hora_inicio_tarea_anterior) + CDate(valor)
                End If
            
            Case 3 ' Segun Receta
                If hora_inicio_tarea_anterior = hora_fin_tarea_anterior Then
                    RstTareasAux("horinitar") = hora_inicio_tarea_anterior
                Else
                    intervalo = Format(Int(HORARR_), "00")
                    intervalo = intervalo & ":" & Format(((HORARR_ * 60) Mod 60), "00")
                    RstTareasAux("horinitar") = CDate(hora_inicio_tarea_anterior) + CDate(intervalo)
                End If
        End Select

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
        fecha_Inicio_Tarea = CDate(Format(fecha_Inicio_Tarea, "dd/mm/yy") & " " & Format(RstTareasAux("horinitar"), "HH:mm")) '+ CDate(Fg2.TextMatrix(A, 16))
        
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
         
        RstTareasAux("idcrdet") = NulosN(ID_CRDET_)
        
        If es_nuevo Then RstTareasAux("activo") = -1
        
        RstTareasAux("idtar") = NulosN(RstTareas_Aux("idtar"))                        ' id de tarea
        RstTareasAux("orden") = NulosN(RstTareas_Aux("orden"))                        ' Orden de la tarea
        RstTareasAux("destar") = NulosC(RstTareas_Aux("destar"))                      ' nombre de la tarea
        RstTareasAux("numper") = Format(NulosN(RstTareas_Aux("numper")), "00")        ' numero de personas para la tarea
        RstTareasAux("aplpor") = Format(NulosN(RstTareas_Aux("aplpor")), FORMAT_CANTIDAD)   ' rendimiento para la cantidad de producto
        
        RstTareasAux("idarea") = NulosN(RstTareas_Aux("idarea"))                        ' Area
        RstTareasAux("desarea") = NulosC(RstTareas_Aux("desarea"))
        RstTareasAux("idtiptrab") = NulosN(RstTareas_Aux("idtiptrab"))                  ' Tipo de Trabajo
        RstTareasAux("destiptrab") = NulosC(RstTareas_Aux("destiptrab"))
        RstTareasAux("idresp") = NulosN(RstTareas_Aux("idresp"))                ' Responsable
        RstTareasAux("nomresp") = NulosC(RstTareas_Aux("nomresp"))
        
        
        RstTareasAux.Update
        
SIGUIENTE:
        RstTareas_Aux.MoveNext
        
        If RstTareas_Aux.EOF = True Then
            Exit For
        End If
    Next B
        
    Agregando = False
End Sub

Private Function calcularHoraFin(IDITEM As Integer, FECHA_DE_INICIO As Date, cantidad As Double) As Date
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

    xTiempo = NulosN(cantidad) / NulosN(RstLinea("frechora"))
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

Private Sub CmdOpciones_Click(Index As Integer)
    Dim xFrm As New sgi2_produccion.produccion
    Dim VISTADISEÑO_ As Boolean
    Dim IDCRDET_ As Double
    Dim Rpta As Integer
        
    VISTADISEÑO_ = False 'Not CalCtrlCronog.Visible
    
    Select Case Index
        Case 0 ' Procesar
            If QueHace = 3 Then Exit Sub
            
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
                    
            If QueHace = 1 Then
                'CalCtrlCronog.Visible = True
                cbfecha.Visible = True
                fg(3).Visible = True
                'CmdOpciones(5).Enabled = True ' Boton cambio de Vista
                CmdOpciones(0).Enabled = False ' Boton Procesar
            End If
'            CalCtrlCronog.ActiveView.ShowDay (CDate(TxtFchIni.valor))
'            CalCtrlCronog.ViewType = xtpCalendarFullWeekView
            
            TxtIdSup.Locked = True
            ComboSemanas.Locked = True
            TxtFchIni.Locked = True
            TxtFchFin.Locked = True
            
            CmdOpciones(1).Enabled = True
            CmdOpciones(2).Enabled = True
            CmdOpciones(3).Enabled = True
            'CmdOpciones(4).Enabled = True
                   
            'If CalCtrlCronog.Visible = True Then CalCtrlCronog.SetFocus
            CmdOpciones_Click 5
        
        Case 1 ' Agregar
            If QueHace = 3 Then Exit Sub
            
            '*************************
            frm(2).Visible = False
            '*************************
            
            If VISTADISEÑO_ Then
                fg(3).Rows = fg(3).Rows + 1
                fg(3).TextMatrix(fg(3).Rows - 1, COLUMNAIDCRDET_) = CORR_
                fg(3).TextMatrix(fg(3).Rows - 1, COLUMNACERRADO_) = ESTADOPENDIENTE_
                fg(3).TextMatrix(fg(3).Rows - 1, COLUMNAPORCENTAJEAPLICADO_) = 100
                fg(3).SetFocus
                If NulosC(cbfecha.Text) <> "TODOS" Then fg(3).TextMatrix(fg(3).Rows - 1, 1) = cbfecha.Text
                fg(3).TopRow = fg(3).Rows - 1
                fg(3).Select fg(3).Rows - 1, 1
                frm(2).Visible = False
                fg(3).TextMatrix(fg(3).Row, COLUMNANUMPROD_) = 0
                CORR_ = CORR_ + 1
            Else
                mostrarFormulario True, False, False, VISTADISEÑO_
            End If
        
        Case 2 ' Modificar
            If QueHace = 3 Then Exit Sub
            mostrarFormulario False, True, False
        
        Case 3 ' Eliminar
            If QueHace = 3 Then Exit Sub
            If VISTADISEÑO_ Then
                ' Eliminar Registros no Pendientes
                If NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNACERRADO_)) <> ESTADOPENDIENTE_ Then
                    MsgBox "No se puede eliminar el registro seleccionado debido al estado en el que se encuentra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    Exit Sub
                End If
                
                IDCRDET_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_))
                
                Rpta = MsgBox("¿Esta seguro de eliminar el registro seleccionado?", _
                                                vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
                If Rpta = vbYes Then
                    ' Se elimina los recordsets relacionados
                    RstProductos.Filter = "id = " & IDCRDET_ & ""
                    RstTareas.Filter = "idcrdet = " & IDCRDET_ & ""
                    RstPersonal.Filter = "idcrdet = " & IDCRDET_ & ""
                    limpiarRST RstProductos, False
                    limpiarRST RstTareas, False
                    limpiarRST RstPersonal, False
                    RstProductos.Filter = adFilterNone
                    RstTareas.Filter = adFilterNone
                    RstPersonal.Filter = adFilterNone
                    
                    fg(3).RemoveItem fg(3).Row
                End If
            Else
                menu2_2_Click
            End If
        
        Case 5 ' Cambiar Vista
            'CalCtrlCronog.Visible = Not CalCtrlCronog.Visible
            VISTADISEÑO_ = False 'Not CalCtrlCronog.Visible
            
            If QueHace = 3 Then
                If VISTADISEÑO_ Then
                    SliderCal.Enabled = False
                Else
                    SliderCal.Enabled = True
                End If
            
            Else
                If VISTADISEÑO_ Then
                    CmdOpciones(2).Enabled = False
                    SliderCal.Enabled = False
                Else
                    CmdOpciones(2).Enabled = True
                    SliderCal.Enabled = True
                End If
            End If
            
            RstProductos.Filter = adFilterNone
            llenarDatos VISTADISEÑO_
            llenarComboFechas
            
    End Select
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
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : LlenarDatos
'* Tipo             : SUB
'* Descripcion      : CARGA LOS DATOS AL CALENDARIO
'* Modificacion     : 15/02/11 JOSE CHACON
'*                      21/04/2011 -> se modifica la referencia "id" de pro_cronogramadet por "idcr"
'*****************************************************************************************************
Sub llenarDatos(Optional DISEÑO_ As Boolean = False, Optional IDCRDETASELEC_ As Double = 0, _
                                                    Optional FECHA_ As String = "TODOS", _
                                                    Optional CERRADO_ As Boolean = False)
    Dim EVENTONUEVO_ As CalendarEvent
    Dim A As Integer
    Dim xRs As New ADODB.Recordset
    Dim FCHEVENTOINI_ As String
    Dim HORAEVENTOINI_ As String
    Dim CONTADOR_ As Integer
    
    Agregando = True
    ' Se llenan los Productos
    If RstProductos.State = 0 Then
        llenarDefinirRST NulosN(RstLis("semana")), False, False, False, True
    End If
    
    RstProductos.Filter = adFilterNone
    
    DEFINIR_RST_TMP xRs, RstProductos
    CARGAR_RST_TMP xRs, RstProductos
    
    xRs.Sort = "fchpro, horpro"
        
    If DISEÑO_ Then
        fg(3).Rows = fg(3).FixedRows
        If xRs.RecordCount = 0 Then Agregando = False: Exit Sub
        xRs.MoveFirst
        FCHEVENTOINI_ = Format(xRs("fchpro"), "dd/mm/yyyy")
        HORAEVENTOINI_ = Format(xRs("horpro"), "HH:mm")
        CONTADOR_ = 1
        For A = 1 To xRs.RecordCount
            ' Se muestran solo los registros de la fecha solicitada
            If NulosC(FECHA_) <> "TODOS" Then
                If Format(xRs("fchpro"), FORMAT_DATE) <> FECHA_ Then GoTo SIGUIENTE
            End If
            fg(3).Rows = fg(3).Rows + 1
            fg(3).TextMatrix(CONTADOR_, COLUMNAFCHPROD_) = Format(xRs("fchpro"), FORMAT_DATE)
            fg(3).TextMatrix(CONTADOR_, COLUMNANUMPROD_) = NulosC(xRs("numprod"))
            fg(3).TextMatrix(CONTADOR_, COLUMNAPRODUCTO_) = NulosC(xRs("descripcion"))
            fg(3).TextMatrix(CONTADOR_, COLUMNARECETA_) = NulosC(xRs("codrec"))
            fg(3).TextMatrix(CONTADOR_, COLUMNAUM_) = NulosC(xRs("abrev"))
            fg(3).TextMatrix(CONTADOR_, COLUMNACANTIDAD_) = Format(NulosN(xRs("cantidad")), FORMAT_CANTIDAD)
            fg(3).TextMatrix(CONTADOR_, COLUMNAENCARGADO_) = NulosC(xRs("nomresp"))
            fg(3).TextMatrix(CONTADOR_, COLUMNALINEA_) = NulosC(xRs("nomlinea"))
            fg(3).TextMatrix(CONTADOR_, COLUMNAHORINI_) = NulosC(Format(xRs("horpro"), FORMAT_HORA_SIN_SEGUNDO))
            fg(3).TextMatrix(CONTADOR_, COLUMNAHORFIN_) = NulosC(Format(xRs("horfin"), FORMAT_HORA_SIN_SEGUNDO))
            fg(3).TextMatrix(CONTADOR_, COLUMNAFCHFIN_) = NulosC(Format(xRs("fchfin"), FORMAT_DATE))
            fg(3).TextMatrix(CONTADOR_, COLUMNANUMOPE_) = Format(NulosN(xRs("numop")), "00")
            fg(3).TextMatrix(CONTADOR_, COLUMNAIDCRDET_) = NulosC(xRs("id"))
            fg(3).TextMatrix(CONTADOR_, COLUMNAIDRECETA_) = NulosC(xRs("idrec"))
            fg(3).TextMatrix(CONTADOR_, COLUMNAIDITEM_) = NulosN(xRs("iditem"))
            fg(3).TextMatrix(CONTADOR_, COLUMNAIDLINEA_) = NulosN(xRs("idlinea"))
            fg(3).TextMatrix(CONTADOR_, COLUMNAIDRESP_) = NulosN(xRs("idresp"))
            fg(3).TextMatrix(CONTADOR_, COLUMNACERRADO_) = NulosN(xRs("estado"))
            fg(3).TextMatrix(CONTADOR_, COLUMNANUMREGPROD_) = Format(NulosC(xRs("numregprod")), "00000000")
            
            '***************************************************************************
            fg(3).TextMatrix(CONTADOR_, COLUMNAPORCENTAJEAPLICADO_) = NulosN(xRs("efic"))
            fg(3).TextMatrix(CONTADOR_, COLUMNAREPROC_) = NulosN(xRs("reproc"))
            '***************************************************************************
            
            If NulosC(xRs("horfin")) = "" And NulosN(xRs("numop")) = 0 Then
                fg(3).TextMatrix(CONTADOR_, COLUMNAPROCESADO_) = ""
            Else
                fg(3).TextMatrix(CONTADOR_, COLUMNAPROCESADO_) = "PROCESADO"
            End If
            CONTADOR_ = CONTADOR_ + 1
SIGUIENTE:
            xRs.MoveNext
        Next A
    Else
'        ' Se pone en blanco el calendario
'        CalCtrlCronog.DataProvider.RemoveAllEvents
'        CalCtrlCronog.Populate
'
'        If xRs.RecordCount = 0 Then Agregando = False: Exit Sub
'        'se crea un evento nuevo de calendario
'        Set EVENTONUEVO_ = CalCtrlCronog.DataProvider.CreateEvent
'
'        'se procede a llenar los detalles del evento
'        xRs.MoveFirst
'        FCHEVENTOINI_ = Format(xRs("fchpro"), "dd/mm/yyyy")
'        HORAEVENTOINI_ = Format(xRs("horpro"), "HH:mm")
'
'        Dim IDEVENTO_ As Long
'        Dim FCHEVENTO_ As Date
'
'        For A = 1 To xRs.RecordCount
'            EVENTONUEVO_.ScheduleID = NulosN(xRs("id"))
'            EVENTONUEVO_.Subject = NulosC(xRs("descripcion"))
'            EVENTONUEVO_.StartTime = Format(xRs("fchpro"), "dd/mm/yyyy") & " " & NulosC(Format(xRs("horpro"), "HH:mm"))
'            EVENTONUEVO_.EndTime = Format(xRs("fchfin"), "dd/mm/yyyy") & " " & NulosC(Format(xRs("horfin"), "HH:mm"))
'            EVENTONUEVO_.Location = NulosC(xRs("numprod"))
'            EVENTONUEVO_.Body = NulosC(xRs("cantidad")) & " " & NulosC(xRs("abrev")) & _
'                                        vbCr + NulosC(Format(xRs("horpro"), "HH:mm")) & " - " _
'                                        & NulosC(Format(xRs("horfin"), "HH:mm"))
'            EVENTONUEVO_.ReminderSoundFile = NulosC(xRs("id"))
'
'            If NulosN(xRs("cerrado")) Then
'                EVENTONUEVO_.Label = 9
'            Else
'                EVENTONUEVO_.Label = 0
'            End If
'
'            'se agrega el evento nuevo al calendario
'            CalCtrlCronog.DataProvider.AddEvent EVENTONUEVO_
'
'            If EVENTONUEVO_.ScheduleID = IDCRDETASELEC_ Then
'                IDEVENTO_ = EVENTONUEVO_.ID
'                FCHEVENTO_ = EVENTONUEVO_.StartTime
'            End If
'
'            xRs.MoveNext
'        Next A
'
'        CalCtrlCronog.Populate
'
'        ' Se posiciona en el primer evento
'        posicionarCelda FCHEVENTOINI_, HORAEVENTOINI_
'        If IDCRDETASELEC_ <> 0 Then
'            seleccionarEvento FCHEVENTO_, IDEVENTO_
'        End If
'
'        Set xRs = Nothing
    End If
    Agregando = False
End Sub

Sub posicionarCelda(FCHBUSCADA_ As String, HORABUSCADA_ As String)
    Dim FCHPOSICION_ As Date
    
    FCHPOSICION_ = CDate(FCHBUSCADA_ & " " & HORABUSCADA_)
    
    With CalCtrlCronog
        .DayView.ScrollV CalCtrlCronog.DayView.GetCellNumber(FCHPOSICION_)
        .DayView.ShowDay FCHPOSICION_
        .ViewType = xtpCalendarFullWeekView
        .ActiveView.SetSelection FCHPOSICION_, FCHPOSICION_, False
        .Populate
    End With
End Sub

Sub seleccionarEvento(FCHBUSCADA_ As Date, IDBUSCADO_ As Long)
    Dim EVENTOAUX_ As CalendarViewEvent
    Dim i As Long
    Dim Ndx As Long
    Dim ENCONTRO_ As Boolean
    
    With CalCtrlCronog
      Ndx = -1
      ENCONTRO_ = False
      
      For i = 0 To .ActiveView.DaysCount - 1
        If .ActiveView(i).Date = DateValue(FCHBUSCADA_) Then
            Ndx = i
            ENCONTRO_ = True
            Exit For
        End If
      Next
    
      If ENCONTRO_ Then
        For Each EVENTOAUX_ In .ActiveView(Ndx)
            If IDBUSCADO_ = EVENTOAUX_.Event.ID Then
                .ActiveView.SelectViewEvent EVENTOAUX_, True
            End If
        Next
      End If
      .Populate
    End With
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
                If CalCtrlCronog.Visible Then CalCtrlCronog.SetFocus
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Function verificarProcesados() As Boolean
    Dim A As Integer
    
    For A = 1 To fg(3).Rows - 1
        If NulosN(fg(3).TextMatrix(A, COLUMNANUMPROD_)) = -1 Then
            verificarProcesados = False
            MsgBox "Hay registros que se han procesado y no se puede salir de la operación", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    Next A
    
    verificarProcesados = True
End Function

Function verificarCampos() As Boolean
    Dim A As Integer
    Dim VERIFICO_ As Boolean
    Dim MENSAJE_ As String
    Dim DISEÑO_ As Boolean
    Dim SUPERVISOR_ As String
    Dim SEMANA_ As String
    Dim PROCESADOS_ As Boolean
    Dim Rpta As Integer
    
    
    DISEÑO_ = Not CalCtrlCronog.Visible
    VERIFICO_ = True
    
    SUPERVISOR_ = NulosC(TxtIdSup.Text)
    SEMANA_ = NulosN(ComboSemanas.Text)
    
    ' Se verifica si hay productos sin procesar
    PROCESADOS_ = True
    If CalCtrlCronog.Visible = True Then GoTo SALIR
    For A = 1 To fg(3).Rows - 1
        If fg(3).TextMatrix(A, COLUMNAPROCESADO_) = "" Then
            PROCESADOS_ = False
            GoTo SIGUIENTE
        End If
    Next A
SIGUIENTE:
    
    If SUPERVISOR_ = "" Then
        MsgBox "No ha especificado un Supervisor para el Cronograma, especifique uno", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        CmdBusSup.SetFocus
        GoTo SALIR
    End If
    
    If SEMANA_ = "" Then
        MsgBox "No ha especificado una fecha para el Cronograma, especifique una", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        VERIFICO_ = False
        ComboSemanas.SetFocus
        GoTo SALIR
    End If
    
    If Not PROCESADOS_ Then
        Rpta = MsgBox("Hay productos sin procesar y no se guardaran, ¿Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
        If Rpta = vbNo Then
            VERIFICO_ = False
        End If
        GoTo SALIR
    End If
SALIR:
    verificarCampos = VERIFICO_
End Function

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
    Dim IDCRREP_ As Double
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
    '***********************************
    Dim RstRep As New ADODB.Recordset
    '***********************************
    Dim RstOrd As New ADODB.Recordset
    Dim RstOrdDet As New ADODB.Recordset
    Dim RstOrdIns As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim xId As Double
    Dim nSQL As String
    Dim pEvent As CalendarEvent
    Dim Events As CalendarEvents
    
    ' VERIFICAMOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
    If Not verificarCampos Then Grabar = False: Exit Function
    
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
            '********************************************************************************
            ' Personal de Produccion
            xCon.Execute "DELETE * FROM pro_cronogramareproceso WHERE idcr  = " & xId & ""
            '********************************************************************************
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
    'Reprocesos
    RST_Busq RstRep, "SELECT TOP 1 * FROM pro_cronogramareproceso", xCon
        
    ' SE LLENA LA CABECERA
    RstCab.AddNew
    RstCab("id") = xId
    RstCab("idsup") = NulosC(TxtIdSup.Text)
    RstCab("fchini") = NulosC(TxtFchIni.valor)
    RstCab("fchfin") = NulosC(TxtFchFin.valor)
    RstCab("semana") = NulosN(ComboSemanas.Text)
    RstCab.Update
    
    Set Events = CalCtrlCronog.DataProvider.GetAllEventsRaw
        
    IDCRDET_ = HallaCodigoTabla("pro_cronogramadet", xCon, "id")
    IDCRPERS_ = HallaCodigoTabla("pro_cronogramapers", xCon, "id")
    IDCRTAR_ = HallaCodigoTabla("pro_cronogramatarea", xCon, "id")
    IDCRREP_ = HallaCodigoTabla("pro_cronogramareproceso", xCon, "id")
    
    RstProductos.Filter = adFilterNone
    RstProductos.MoveFirst
    For A = 1 To RstProductos.RecordCount
        Dim IDCRDETAUX_ As Double
        IDCRDETAUX_ = NulosN(RstProductos("id"))
        
        RstDet.AddNew
        RstDet("id") = IDCRDET_
        RstDet("idcr") = xId
        RstDet("fchpro") = NulosC(Format(RstProductos("fchpro"), "dd/mm/yyyy"))
        RstDet("fchfin") = NulosC(Format(RstProductos("fchfin"), "dd/mm/yyyy"))
        RstDet("horpro") = NulosC(Format(RstProductos("horpro"), "HH:mm"))
        RstDet("horfin") = NulosC(Format(RstProductos("horfin"), "HH:mm"))
        RstDet("iditem") = NulosN(RstProductos("iditem"))
        RstDet("idrec") = NulosN(RstProductos("idrec"))
        RstDet("cantidad") = NulosN(RstProductos("cantidad"))
        RstDet("numprod") = NulosC(RstProductos("numprod"))
        RstDet("idresp") = NulosN(RstProductos("idresp"))
        RstDet("idlinea") = NulosN(RstProductos("idlinea"))
        RstDet("numop") = NulosN(RstProductos("numop"))
        RstDet("cerrado") = NulosN(RstProductos("cerrado"))
        RstDet("idprocorr") = NulosN(RstProductos("idprocorr"))
        RstDet("estado") = NulosN(RstProductos("estado"))
        RstDet("efic") = NulosN(RstProductos("efic"))
        RstDet("reproc") = NulosN(RstProductos("reproc"))
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
                RstTar("idpro") = NulosN(RstProductos("iditem"))
                RstTar("fchpro") = NulosC(RstProductos("fchpro"))
                RstTar("idtar") = NulosN(RstTareas("idtar"))
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
                RstTar("idtiptrab") = NulosN(RstTareas("idtiptrab"))
                RstTar("idarea") = NulosN(RstTareas("idarea"))
                RstTar("idresp") = NulosN(RstTareas("idresp"))
                RstTar.Update
                
                IDCRTAR_ = IDCRTAR_ + 1
                RstTareas.MoveNext
            Next B
        End If
                
        RstReproc.Filter = "idcrdet = " & IDCRDETAUX_ & ""
        If RstReproc.RecordCount <> 0 Then
            RstReproc.MoveFirst
            For B = 1 To RstReproc.RecordCount
                RstRep.AddNew
                RstRep("id") = IDCRREP_
                RstRep("idcr") = xId
                RstRep("idcrdet") = IDCRDET_
                RstRep("idlotedet") = NulosN(RstReproc("idlotedet"))
                RstRep("cantidad") = NulosN(RstReproc("cantidad"))
                RstRep.Update
                
                IDCRREP_ = IDCRREP_ + 1
                RstReproc.MoveNext
            Next B
        End If
        
        IDCRDET_ = IDCRDET_ + 1
        
        RstProductos.MoveNext
    Next A
    
    ' Grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
   
    xCon.CommitTrans
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

Private Sub DTPHoras_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub fg_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Index
        Case 0 ' Tareas Calendario
            If cbEstado.ItemData(cbEstado.ListIndex) > ESTADOPENDIENTE_ Then Cancel = True: Exit Sub
            
            Select Case Col
                Case 2 To 9
                    Cancel = True
                    
            End Select
        
        Case 1 ' Personal Calendario
            If cbEstado.ItemData(cbEstado.ListIndex) > ESTADOPENDIENTE_ Then Cancel = True: Exit Sub
            
            Select Case Col
                Case 2 To 3
                    Cancel = True
                    
            End Select
        
        Case 3 ' fg de Productos Diseño
            If fg(3).Row < fg(3).FixedRows Then Exit Sub
            If NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNACERRADO_)) > ESTADOPENDIENTE_ Then Cancel = True: Exit Sub
            
            Select Case Col
                Case COLUMNACERRADO_, COLUMNANUMPROD_, COLUMNAHORFIN_, COLUMNAFCHFIN_, _
                                            COLUMNAPROCESADO_, COLUMNAUM_, COLUMNANUMREGPROD_
                    Cancel = True
                    
            End Select
        
        Case 5 ' Tareas Diseño
            If NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNACERRADO_)) > ESTADOPENDIENTE_ Then Cancel = True: Exit Sub
            
            Select Case Col
                Case 2 To 9
                    Cancel = True
                    
            End Select
            
        Case 4 ' Personal Diseño
            If NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNACERRADO_)) > ESTADOPENDIENTE_ Then Cancel = True: Exit Sub
            
            Select Case Col
                Case 2 To 3
                    Cancel = True
                    
            End Select
            
    End Select
End Sub

Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim xCampos(2, 4) As String
    Dim xRs As New ADODB.Recordset
    Dim nTitulo As String
    Dim IDCRDET_ As Double
    Dim DISEÑO_ As Boolean
    
    If CalCtrlCronog.Visible Then
        DISEÑO_ = False
    Else
        DISEÑO_ = True
    End If
    
    Select Case Index
        Case 0 ' fg de Tareas
            If QueHace = 3 Then Exit Sub
            
            If DISEÑO_ Then
                IDCRDET_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_))
            Else
                IDCRDET_ = NulosN(LblIdCrDet.Caption)
            End If
            
            Select Case Col
                Case 13 ' Area
                    xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
                    xCampos(1, 0) = "Id":           xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
                    
                    nTitulo = "Buscando Area"
                    
                    cSQL = "SELECT mae_area.id, mae_area.descripcion AS nombre, mae_area.id AS cod, mae_area.id AS idarea, pro_emp.id AS idper, pla_empleados.id AS idemp, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS encargado, pla_empleados.numdoc " _
                        + vbCr + "FROM (((pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) RIGHT JOIN pro_area ON pro_emp.id = pro_area.idper) INNER JOIN mae_area ON pro_area.idarea = mae_area.id) LEFT JOIN pro_areadet ON pro_area.id = pro_areadet.idar " _
                        + vbCr + "WHERE (((pro_areadet.idtar)=" & NulosN(fg(Index).TextMatrix(Row, 11)) & ")); "
                    
                    CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                                    "nombre", "nombre", Principio, ""
                                                                  
                    If xRs.State = 0 Then Exit Sub
                    If xRs.RecordCount = 0 Then Exit Sub
                    
                    fg(Index).TextMatrix(Row, Col) = NulosC(xRs("nombre"))
                    fg(Index).TextMatrix(Row, 16) = NulosN(xRs("id"))
                    
                    If RstTareasAux.State = 0 Then Exit Sub
                    RstTareasAux.Filter = "idcrdet = " & IDCRDET_ & " And idtar = " & NulosN(fg(Index).TextMatrix(fg(Index).Row, 11)) & ""
                    If RstTareasAux.RecordCount = 0 Then Exit Sub
                    
                    RstTareasAux("desarea") = NulosC(xRs("nombre"))
                    RstTareasAux("idarea") = NulosC(xRs("id"))
                    
                    ' Se Agrega responsable de Area
                    fg(0).TextMatrix(Row, 15) = NulosC(xRs("encargado"))  ' Responsable
                    fg(0).TextMatrix(Row, 18) = NulosN(xRs("idemp"))  ' idresponsable
                    
                    If RstTareasAux.State = 0 Then Exit Sub
                    RstTareasAux.Filter = "idcrdet = " & IDCRDET_ & " And idtar = " & NulosN(fg(Index).TextMatrix(fg(Index).Row, 11)) & ""
                    If RstTareasAux.RecordCount = 0 Then Exit Sub
                    
                    RstTareasAux("nomresp") = NulosC(xRs("encargado"))
                    RstTareasAux("idresp") = NulosC(xRs("idemp"))
                    
                Case 14 ' Tipo de Trabajo
                    'ReDim xCampos(2, 4) As String
                    xCampos(0, 0) = "Id":           xCampos(0, 1) = "id":           xCampos(0, 2) = "500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
                    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "2500":    xCampos(1, 3) = "C":   xCampos(1, 4) = "C"
                    
                    nTitulo = "Buscando Tipos de Trabajo"
                    
                    cSQL = "SELECT * FROM pro_tiptrab"
                    CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                                    "descripcion", "descripcion", Principio, ""
                    
                    If xRs.State = 0 Then Exit Sub
                    If xRs.RecordCount = 0 Then Exit Sub
                    
                    fg(Index).TextMatrix(Row, Col) = NulosC(xRs("descripcion"))
                    fg(Index).TextMatrix(Row, 17) = NulosN(xRs("id"))
                    
                    If RstTareasAux.State = 0 Then Exit Sub
                    RstTareasAux.Filter = "idcrdet = " & IDCRDET_ & " And idtar = " & NulosN(fg(Index).TextMatrix(fg(Index).Row, 11)) & ""
                    If RstTareasAux.RecordCount = 0 Then Exit Sub
                    
                    RstTareasAux("destiptrab") = NulosC(xRs("descripcion"))
                    RstTareasAux("idtiptrab") = NulosC(xRs("id"))
                    
                Case 15 ' Responsable de Tarea
                    'ReDim xCampos(2, 4) As String
                    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
                    xCampos(0, 0) = "Apellidos y Nombres":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
                    xCampos(1, 0) = "Codigo":                xCampos(1, 1) = "id":        xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
                            
                    cSQL = "SELECT pro_emp.idemp, pro_emp.sup, pro_emp.prog, pro_emp.res, pla_empleados.nombre " _
                        + vbCr + "FROM (pro_emp LEFT JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id " _
                        + vbCr + "Where (((pro_empdet.idfun) = 3)) " _
                        + vbCr + "GROUP BY pro_emp.idemp, pro_emp.sup, pro_emp.prog, pro_emp.res, pla_empleados.nombre " _
                        + vbCr + "Having (((pla_empleados.nombre) Is Not Null)) " _
                        + vbCr + "ORDER BY pla_empleados.nombre;"
                        
                    nTitulo = "Buscando Personal Responsable"
                            
                    CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
                    
                    If xRs.State = 0 Then Exit Sub
                    If xRs.RecordCount = 0 Then Exit Sub
                    
                    fg(0).TextMatrix(Row, Col) = NulosC(xRs("nombre"))  ' Responsable
                    fg(0).TextMatrix(Row, 18) = NulosN(xRs("idemp"))  ' idresponsable
                    
                    If RstTareasAux.State = 0 Then Exit Sub
                    RstTareasAux.Filter = "idcrdet = " & IDCRDET_ & " And idtar = " & NulosN(fg(Index).TextMatrix(fg(Index).Row, 11)) & ""
                    If RstTareasAux.RecordCount = 0 Then Exit Sub
                    
                    RstTareasAux("nomresp") = NulosC(xRs("nombre"))
                    RstTareasAux("idresp") = NulosC(xRs("idemp"))
                    
            End Select
            
        Case 3 ' fg de vista de Mantenimiento
            If CalCtrlCronog.Visible = True Then Exit Sub
            If NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNACERRADO_)) <> ESTADOPENDIENTE_ Then Exit Sub
            Select Case Col
                Case COLUMNAPRODUCTO_: ' Producto
                    agregarCampos True, False, False, False, False, True
                    fg(3).Select fg(3).Row, COLUMNACANTIDAD_
                Case COLUMNARECETA_:  ' Receta
                    agregarCampos False, False, False, True, False, True
                    fg(3).Select fg(3).Row, COLUMNACANTIDAD_
                Case COLUMNAENCARGADO_:  ' Encargado
                    agregarCampos False, False, True, False, False, True
                    fg(3).Select fg(3).Row, COLUMNAHORINI_
                Case COLUMNALINEA_:  ' Linea
                    agregarCampos False, False, False, False, True, True
                    fg(3).Select fg(3).Row, COLUMNAHORINI_
            End Select
            
        Case 5 ' fg de Tareas Mantenimiento
            If QueHace = 3 Then Exit Sub
            IDCRDET_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_))
            Select Case Col
                Case 13 ' Area
                    'ReDim xCampos(2, 4) As String
                    xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
                    xCampos(1, 0) = "Id":           xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
                    
                    nTitulo = "Buscando Area"
                    
                    cSQL = "SELECT mae_area.id, mae_area.descripcion AS nombre, mae_area.id AS cod, mae_area.id AS idarea, pro_emp.id AS idper, pla_empleados.id AS idemp, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS encargado, pla_empleados.numdoc " _
                        + vbCr + " FROM ((pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) RIGHT JOIN pro_area ON pro_emp.id = pro_area.idper) INNER JOIN mae_area ON pro_area.idarea = mae_area.id "
                    CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                                    "nombre", "nombre", Principio, ""
                                                                  
                    If xRs.State = 0 Then Exit Sub
                    If xRs.RecordCount = 0 Then Exit Sub
                    
                    fg(Index).TextMatrix(Row, Col) = NulosC(xRs("nombre"))
                    fg(Index).TextMatrix(Row, 16) = NulosN(xRs("id"))
                    
                    If RstTareas.State = 0 Or RstTareas.State = 0 Then Exit Sub
                    RstTareas.Filter = "idcrdet = " & IDCRDET_ & " And idtar = " & NulosN(fg(Index).TextMatrix(fg(Index).Row, 11)) & ""
                    RstTareasAux.Filter = "idcrdet = " & IDCRDET_ & " And idtar = " & NulosN(fg(Index).TextMatrix(fg(Index).Row, 11)) & ""
                    If RstTareas.RecordCount = 0 Or RstTareasAux.RecordCount = 0 Then Exit Sub
                    
                    RstTareasAux("desarea") = NulosC(xRs("nombre"))
                    RstTareasAux("idarea") = NulosC(xRs("id"))
                    RstTareasAux.Update
                    
                    RstTareas("desarea") = NulosC(xRs("nombre"))
                    RstTareas("idarea") = NulosC(xRs("id"))
                    RstTareas.Update
                    
                Case 14 ' Tipo de Trabajo
                    'ReDim xCampos(2, 4) As String
                    xCampos(0, 0) = "Id":           xCampos(0, 1) = "id":           xCampos(0, 2) = "500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
                    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "2500":    xCampos(1, 3) = "C":   xCampos(1, 4) = "C"
                    
                    nTitulo = "Buscando Tipos de Trabajo"
                    
                    cSQL = "SELECT * FROM pro_tiptrab"
                    CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                                    "descripcion", "descripcion", Principio, ""
                                                                    
                    
                    If xRs.State = 0 Then Exit Sub
                    If xRs.RecordCount = 0 Then Exit Sub
                    
                    fg(Index).TextMatrix(Row, Col) = NulosC(xRs("descripcion"))
                    fg(Index).TextMatrix(Row, 17) = NulosN(xRs("id"))
                    
                    If RstTareas.State = 0 Or RstTareasAux.State = 0 Then Exit Sub
                    RstTareas.Filter = "idcrdet = " & IDCRDET_ & " And idtar = " & NulosN(fg(Index).TextMatrix(fg(Index).Row, 11)) & ""
                    RstTareasAux.Filter = "idcrdet = " & IDCRDET_ & " And idtar = " & NulosN(fg(Index).TextMatrix(fg(Index).Row, 11)) & ""
                    If RstTareas.RecordCount = 0 Or RstTareasAux.RecordCount = 0 Then Exit Sub
                    
                    RstTareasAux("destiptrab") = NulosC(xRs("descripcion"))
                    RstTareasAux("idtiptrab") = NulosC(xRs("id"))
                    RstTareasAux.Update
                    
                    RstTareas("destiptrab") = NulosC(xRs("descripcion"))
                    RstTareas("idtiptrab") = NulosC(xRs("id"))
                    RstTareas.Update
                    
                Case 15 ' Forma de Pago
                    'ReDim xCampos(2, 4) As String
                    xCampos(0, 0) = "Id":           xCampos(0, 1) = "id":           xCampos(0, 2) = "500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
                    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "2500":    xCampos(1, 3) = "C":   xCampos(1, 4) = "C"
                    
                    nTitulo = "Buscando Formas de Pago"
                    
                    cSQL = "SELECT * FROM pro_formapag"
                    CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), nTitulo, _
                                                                    "descripcion", "descripcion", Principio, ""
                                                                    
                    If xRs.State = 0 Then Exit Sub
                    If xRs.RecordCount = 0 Then Exit Sub
                    
                    fg(Index).TextMatrix(Row, Col) = NulosC(xRs("descripcion"))
                    fg(Index).TextMatrix(Row, 18) = NulosN(xRs("id"))
                    
                    If RstTareas.State = 0 Or RstTareasAux.State = 0 Then Exit Sub
                    RstTareas.Filter = "idcrdet = " & IDCRDET_ & " And idtar = " & NulosN(fg(Index).TextMatrix(fg(Index).Row, 11)) & ""
                    RstTareasAux.Filter = "idcrdet = " & IDCRDET_ & " And idtar = " & NulosN(fg(Index).TextMatrix(fg(Index).Row, 11)) & ""
                    If RstTareas.RecordCount = 0 Or RstTareasAux.RecordCount = 0 Then Exit Sub
                    
                    RstTareasAux("desformapag") = NulosC(xRs("descripcion"))
                    RstTareasAux("idformapag") = NulosC(xRs("id"))
                    RstTareasAux.Update
                    
                    RstTareas("desformapag") = NulosC(xRs("descripcion"))
                    RstTareas("idformapag") = NulosC(xRs("id"))
                    RstTareas.Update
                    
            End Select
        
    End Select
End Sub

Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim Rpta As Integer
    Dim IDCRDET_ As Double
    Dim NINGUNERROR_ As Boolean
    Dim MENSAJE_ As String
    
    If Index = 3 Then
        If Agregando Then Exit Sub
        If NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNACERRADO_)) = 3 Then Exit Sub
        
        Select Case Col
            Case COLUMNACANTIDAD_
                fg(3).TextMatrix(Row, Col) = Format(fg(3).TextMatrix(Row, Col), FORMAT_CANTIDAD)
                ' Se quita el estatus de Procesado
                fg(3).TextMatrix(Row, COLUMNAHORFIN_) = ""
                fg(3).TextMatrix(Row, COLUMNAFCHFIN_) = ""
                fg(3).TextMatrix(Row, COLUMNANUMOPE_) = ""
                fg(3).TextMatrix(Row, COLUMNAPROCESADO_) = ""
                fg(3).Select Row, 1, Row, fg(3).Cols - 1
                fg(3).FillStyle = flexFillRepeat
                fg(3).CellBackColor = &HB9B9FF
                fg(3).Select Row, Col
                
            Case COLUMNAHORINI_
                If IsDate(fg(3).TextMatrix(Row, Col)) Then
                    fg(3).TextMatrix(Row, Col) = Format(fg(3).TextMatrix(Row, Col), FORMAT_HORA_SIN_SEGUNDO)
                Else
                    MsgBox "Ingrese una Hora adecuada", vbInformation, xTitulo
                    fg(3).TextMatrix(Row, Col) = ""
                    fg(3).Select Row, Col
                    Exit Sub
                End If
                
                ' Se quita el estatus de Procesado
                fg(3).TextMatrix(Row, COLUMNAHORFIN_) = ""
                fg(3).TextMatrix(Row, COLUMNAFCHFIN_) = ""
                fg(3).TextMatrix(Row, COLUMNANUMOPE_) = ""
                fg(3).TextMatrix(Row, COLUMNAPROCESADO_) = ""
                fg(3).Select Row, 1, Row, fg(3).Cols - 1
                fg(3).FillStyle = flexFillRepeat
                fg(3).CellBackColor = &HB9B9FF
                fg(3).Select Row, Col
                
                ' Si todos los campos estan correctos
                If frm(2).Visible Then Exit Sub
                If NulosN(fg(3).TextMatrix(Row, COLUMNAIDITEM_)) <> 0 _
                            And NulosN(fg(3).TextMatrix(Row, COLUMNAIDRECETA_)) <> 0 _
                            And NulosN(fg(3).TextMatrix(Row, COLUMNAIDLINEA_)) <> 0 _
                            And NulosN(fg(3).TextMatrix(Row, COLUMNACANTIDAD_)) <> 0 _
                            And NulosC(fg(3).TextMatrix(Row, COLUMNAHORINI_)) <> "" Then
                           
                    bloquearControles NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_)), True
                    mostrarFormulario True, False, False, True
                    Cmd(2).SetFocus
                End If
                
            Case COLUMNAFCHPROD_
                If Agregando Then Exit Sub
                If Not IsDate(fg(3).TextMatrix(Row, Col)) Then
                    MsgBox "La fecha ingresada es incorrecta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    fg(3).TextMatrix(Row, Col) = ""
                    fg(3).Select Row, Col
                    Exit Sub
                End If
                                
                If cbfecha.Text = "TODOS" Then
                    If (CDate(fg(3).TextMatrix(Row, Col)) < CDate(TxtFchIni.valor)) Or _
                                        (CDate(fg(3).TextMatrix(Row, Col)) > CDate(TxtFchFin.valor)) Then
                        MsgBox "La fecha ingresada es incorrecta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                        fg(3).TextMatrix(Row, Col) = ""
                        fg(3).Select Row, Col
                    End If
                Else
                    If (fg(3).TextMatrix(Row, Col) <> Format(cbfecha.Text, FORMAT_DATE)) Then
                        MsgBox "La fecha ingresada es incorrecta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                        fg(3).TextMatrix(Row, Col) = ""
                    End If
                End If
                            
            Case COLUMNAIDITEM_, COLUMNAIDLINEA_, COLUMNAIDRECETA_, COLUMNAIDRESP_
                ' Se quita el estatus de Procesado
                fg(3).TextMatrix(Row, COLUMNAPROCESADO_) = ""
                fg(3).Select Row, 1, Row, fg(3).Cols - 1
                fg(3).FillStyle = flexFillRepeat
                fg(3).CellBackColor = &HB9B9FF
                fg(3).Select Row, Col
                          
        End Select
    End If
End Sub

Private Sub Fg_Click(Index As Integer)
    Dim Rpta As Integer
    Dim ESPECIAL_ As Boolean
    
    If QueHace = 3 Then Exit Sub
    
    ESPECIAL_ = False
    
    Select Case Index
        Case 0 ' fg de Tareas
            Dim DISEÑO_ As Boolean
            Dim IDCRDET_ As Double
            
            ' Se verifica si se esta o no en vista de diseño
            DISEÑO_ = False 'Not CalCtrlCronog.Visible
            
            If DISEÑO_ Then
                IDCRDET_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_))
            Else
                'IDCRDET_ = NulosN(LblIdCrDet.Caption)
            End If
            
            If fg(Index).Row < fg(Index).FixedRows Then Exit Sub
            If fg(Index).Col <> 1 Then Exit Sub
            
            If fg(Index).TextMatrix(fg(Index).Row, 1) = 0 Then ' Si se deselecciono
                'xTitulo = "Cambio en el estado de Tarea"
                RstPersonalAux.Filter = "idcrdet = " & IDCRDET_ & " And idtar = " & NulosN(fg(Index).TextMatrix(fg(Index).Row, 11)) & ""
                If RstPersonalAux.RecordCount > 0 Then
                    Rpta = MsgBox("¿Se eliminara el Personal relacionado a esta Tarea; desea continuar?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
                    If Rpta = vbNo Then
                        ' Se selecciona de nuevo la tarea
                        fg(Index).TextMatrix(fg(Index).Row, 1) = -1
                        Exit Sub
                    End If
                    fg(1).Rows = fg(Index).FixedRows
                End If
                
                ' Se filtra el Personal de la Tarea y se elimina
                If fg(Index).Row > 1 Then ESPECIAL_ = True
                RstPersonalAux.Filter = "idcrdet = " & IDCRDET_ & " And idtar = " & NulosN(fg(Index).TextMatrix(fg(Index).Row, 11)) & ""
                limpiarRST RstPersonalAux, False
                
'                If DISEÑO_ Then
'                    fg(4).Rows = fg(4).FixedRows
'                End If
                ' Se filtra la Tarea y se actualiza su estado
                RstTareasAux.Filter = "idcrdet = " & IDCRDET_ & " And idtar = " & NulosN(fg(Index).TextMatrix(fg(Index).Row, 11)) & ""
                If RstTareasAux.RecordCount > 0 Then RstTareasAux("activo") = False
                ' Se modifica las tareas
                limpiarTarea IDCRDET_, NulosN(fg(Index).TextMatrix(fg(Index).Row, 11))
            Else
                ' Se filtra el Personal de la Tarea y se elimina
                RstTareasAux.Filter = "idcrdet = " & IDCRDET_ & " And idtar = " & NulosN(fg(Index).TextMatrix(fg(Index).Row, 11)) & ""
                If RstTareasAux.RecordCount > 0 Then RstTareasAux("Activo") = True
            End If
            
        Case 1 ' fg de Personal
        
        Case 2 ' fg de Ranking de Personal
            Dim A As Integer
            Dim contador As Integer
            
            If frm(4).Visible = False Then Exit Sub
            
            contador = 0
            For A = 1 To fg(2).Rows - 1
                If fg(2).TextMatrix(A, 1) = -1 Then contador = contador + 1
            Next A
            
            LbNumSel.Caption = Format(contador, "000")
    End Select
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
        If .RecordCount = 0 Then RENDIMIENTO_ = 1: GoTo SALIR_
        .MoveFirst
        RENDIMIENTO_ = 1
        For A = 1 To .RecordCount
            xRs.Filter = "idtar = " & NulosN(.Fields("idtar"))
            If xRs.RecordCount = 0 Then GoTo SIGUIENTE_
            RENDIMIENTO_ = RENDIMIENTO_ * (NulosN(xRs("rdmto")) / 100)
            .MoveNext
SIGUIENTE_:
        Next A
    End With
SALIR_:
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

Private Sub Fg_DblClick(Index As Integer)
    If Index = 3 Then
        ' Se cargan las tareas
        Agregando = True
        If NulosC(fg(3).TextMatrix(fg(3).Row, COLUMNAPROCESADO_)) = "" _
                            And NulosC(fg(3).TextMatrix(fg(3).Row, COLUMNAHORFIN_)) = "" _
                            And NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNANUMOPE_)) = 0 Then
            mostrarFormulario True, False, False, True
        Else
            mostrarFormulario False, True, False, True
        End If
        Agregando = False
        bloquearControles fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_), True
    End If
End Sub

Private Sub Fg_EnterCell(Index As Integer)
    If Agregando Then Exit Sub
        
    If QueHace = 3 Then
        fg(Index).Editable = flexEDNone
        Exit Sub
    End If
    
    fg(Index).Editable = flexEDKbdMouse
End Sub

Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Index = 3 Then
        Select Case Col
            Case COLUMNAPRODUCTO_, COLUMNARECETA_, COLUMNAENCARGADO_
                KeyAscii = 0
        End Select
    End If
End Sub

Private Sub Fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> 3 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    If Button = 2 Then
        PopupMenu menu4
    End If
End Sub

Private Sub totalizarPersonal()
    Dim DISEÑO_ As Boolean
    Dim IDCRDET_ As Double
    Dim NUMOP_ As Double
    Dim NUMOPSEL_ As Double
    
    If CalCtrlCronog.Visible Then DISEÑO_ = False Else DISEÑO_ = True
    If DISEÑO_ Then
        IDCRDET_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_))
    Else
        IDCRDET_ = NulosN(LblIdCrDet.Caption)
    End If
    
    ' Se actualiza el detalle de los trabajadores
    ' Total
    RstPersonalAux.Filter = "idcrdet = " & IDCRDET_
    NUMOP_ = NulosN(lblntrabtot.Caption)
    NUMOPSEL_ = RstPersonalAux.RecordCount
    
    lblNumOpe.Caption = Format(NUMOPSEL_, "00") & " de " & Format(NUMOP_, "00")
    ' Parcial
    lblntrab.Caption = NulosN(fg(0).TextMatrix(fg(0).Row, 6))
    LblDetTrab.Caption = Format(fg(1).Rows - 1, "00") & " de " & Format(NulosN(lblntrab.Caption), "00")
End Sub

Private Sub fg_RowColChange(Index As Integer)
    Dim DISEÑO_ As Boolean
    
    DISEÑO_ = False 'Not CalCtrlCronog.Visible
    
    Select Case Index
        Case 0 ' fg Tareas
            If Agregando Then Exit Sub
            ' Se carga el Personal
            pCargarDatos fg(1), True, False, False, False, False, DISEÑO_
            totalizarPersonal
            
        Case 3 ' fg Diseño
'            If frm(2).Visible = False Then Exit Sub
'            If Agregando Then Exit Sub
'
'            ' Se carga las Tareas
'            If NulosC(fg(3).TextMatrix(fg(3).Row, COLUMNAPROCESADO_)) = "" Then
'                mostrarFormulario True, False, False, True
'            Else
'                mostrarFormulario False, True, False, True
'            End If
        
        Case 5 ' fg Diseño Tareas
    End Select
End Sub

Private Sub Form_Activate()    '
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
    
    Agregando = True
    
    TxtIdSup.Text = RstLis("idsup")
    TxtIdSup_Validate True
    
    TxtFchIni.valor = RstLis("fchini")
    TxtFchFin.valor = RstLis("fchfin")
    
'    CalCtrlCronog.ActiveView.ShowDay (CDate(TxtFchIni.valor))
'    CalCtrlCronog.ViewType = xtpCalendarFullWeekView
    
    'llenarCargarEstado
    llenarEstado False, , fg(3), COLUMNACERRADO_
    
    CentrarFrm frm(3)
    frm(3).Visible = True
    frm(3).Refresh
    LblProg.Caption = "PROCESANDO PRODUCTOS"
    llenarDatos
    
    frm(3).Refresh
    LblProg.Caption = "PROCESANDO TAREAS"
    llenarDefinirRST NulosN(RstLis("semana")), True, False ' Tareas
    
    frm(3).Refresh
    LblProg.Caption = "PROCESANDO PERSONAL"
    llenarDefinirRST NulosN(RstLis("semana")), False, True ' Personal
    
    frm(3).Refresh
    LblProg.Caption = "PROCESANDO REPROCESOS"
    llenarDefinirRST NulosN(RstLis("semana")), False, False, , , , True ' Reproceso
    
    frm(3).Visible = False
    
    CARGO_ = True ' Se define como cargo
    CmdOpciones_Click 5
    
    Agregando = False
End Sub

Private Sub llenarDefinirRST(SEMANA_ As Double, TAREAS_ As Boolean, PERSONAL_ As Boolean, _
                                Optional RECETA_ As Boolean = False, Optional PRODUCTOS_ As Boolean = False, _
                                Optional MATERIAPRIMA_ As Boolean = False, Optional REPROCESO_ As Boolean = False)
    Dim xRs As New ADODB.Recordset
    
    If TAREAS_ Then
        ' Se llena las Tareas
        cSQL = "SELECT pro_cronogramatarea.idtar, pro_cronogramatarea.orden, pro_tareas.descripcion AS destar, pro_cronogramatarea.idcr, pro_cronogramatarea.idcrdet, pro_cronogramatarea.idlinea, pro_cronogramatarea.activo, pro_cronogramatarea.cantproc, pro_cronogramatarea.numper, pro_cronogramatarea.horinitar, pro_cronogramatarea.horfintar, pro_cronogramatarea.durtar, pro_cronogramatarea.fchini, pro_cronogramatarea.fchfin, pro_cronogramatarea.aplpor, pro_cronogramatarea.idresp, pla_empleados.nombre AS nomresp, pro_cronogramatarea.idarea, pro_cronogramatarea.idtiptrab, pro_cronogramatarea.idformapag, mae_area.descripcion AS desarea, pro_tiptrab.descripcion AS destiptrab, pro_formapag.descripcion AS desformapag " _
            + vbCr + "FROM ((((pro_cronograma RIGHT JOIN ((pro_cronogramatarea LEFT JOIN pro_tareas ON pro_cronogramatarea.idtar = pro_tareas.id) LEFT JOIN alm_inventario ON pro_cronogramatarea.idpro = alm_inventario.id) ON pro_cronograma.id = pro_cronogramatarea.idcr) LEFT JOIN pla_empleados ON pro_cronogramatarea.idresp = pla_empleados.id) LEFT JOIN mae_area ON pro_cronogramatarea.idarea = mae_area.id) LEFT JOIN pro_formapag ON pro_cronogramatarea.idformapag = pro_formapag.id) LEFT JOIN pro_tiptrab ON pro_cronogramatarea.idtiptrab = pro_tiptrab.id " _
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
        Set xRs = Nothing
    End If
    
    If PERSONAL_ Then
        ' Se llena al personal
        cSQL = "SELECT pro_cronogramapers.idper, pla_empleados.nombre, pla_empleados.numdoc, pro_cronogramapers.idcr, pro_cronogramapers.idcrdet, pro_cronogramapers.idtar, pro_cronogramapers.activo, pro_tareas.descripcion AS destar " _
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
        cSQL = "SELECT pro_cronogramadet.*, alm_inventario.descripcion, mae_unidades.abrev, pro_receta.codrec, pla_empleados.nombre AS nomresp, pro_linea.descripcion AS nomlinea, pro_producciondet.numparte AS numregprod " _
            + vbCr + "FROM ((((pro_cronograma LEFT JOIN ((pro_cronogramadet LEFT JOIN alm_inventario ON pro_cronogramadet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON alm_inventario.idunimed = mae_unidades.id) ON pro_cronograma.id = pro_cronogramadet.idcr) LEFT JOIN pro_receta ON pro_cronogramadet.idrec = pro_receta.id) LEFT JOIN pla_empleados ON pro_cronogramadet.idresp = pla_empleados.id) LEFT JOIN pro_linea ON pro_cronogramadet.idlinea = pro_linea.id) LEFT JOIN pro_producciondet ON pro_cronogramadet.idprocorr = pro_producciondet.corr " _
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
    
    If REPROCESO_ Then
        cSQL = "SELECT pro_cronogramareproceso.idcr, pro_cronogramareproceso.idcrdet, pro_cronogramareproceso.idlotedet, alm_inventariolote.descripcion AS deslote, pro_cronogramareproceso.cantidad, alm_inventariolotedet.idalm, alm_almacenes.descripcion AS desalm " _
            + vbCr + "FROM (pro_cronograma RIGHT JOIN ((pro_cronogramareproceso LEFT JOIN alm_inventariolotedet ON pro_cronogramareproceso.idlotedet = alm_inventariolotedet.id) LEFT JOIN alm_inventariolote ON alm_inventariolotedet.idlote = alm_inventariolote.id) ON pro_cronograma.id = pro_cronogramareproceso.idcr) LEFT JOIN alm_almacenes ON alm_inventariolotedet.idalm = alm_almacenes.id " _
            + vbCr + "WHERE (((pro_cronograma.semana)=" & SEMANA_ & "));"
        
        RST_Busq xRs, cSQL, xCon
        
        If RstReproc.State = 0 Then DEFINIR_RST_TMP RstReproc, xRs
        limpiarRST RstReproc
            
        CARGAR_RST_TMP RstReproc, xRs
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
    habilitar CmdOpciones, True
    Bloquea
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
        llenarDefinirRST 99999, False, False, False, False, False, True ' Reproceso
    End If
    
    Agregando = True
    llenarEstado False, , fg(3), COLUMNACERRADO_
    llenarEstado True
    Agregando = False
    
    ComboSemanas.Locked = False
    CmdOpciones(0).Enabled = True
    CmdBusSup.SetFocus
    ARRASTRANDO_ = False
End Sub

Private Function verificarCambioEstado(IDPROCORR_ As Double, ByRef MENSAJE_ As String) As Boolean
    Dim xRs As New ADODB.Recordset
    
    ' Buscando Registros de Produccion
    cSQL = "SELECT pro_producciondet.corr, pro_producciondet.estado " _
        + vbCr + "FROM pro_producciondet " _
        + vbCr + "WHERE (((pro_producciondet.corr)=" & IDPROCORR_ & ") AND ((pro_producciondet.estado)>=2));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    MENSAJE_ = "Registros de Produccción"
    
    If xRs.State = 0 Then verificarCambioEstado = False: GoTo SALIR_
    If xRs.RecordCount > 0 Then verificarCambioEstado = False: GoTo SALIR_
    
    ' Buscando Registros de Solicitud de Materiales
    cSQL = "SELECT pro_ordenproddet.idprocorr, pro_ordenproddet.estado " _
        + vbCr + "FROM pro_ordenproddet " _
        + vbCr + "WHERE (((pro_ordenproddet.idprocorr)=" & IDPROCORR_ & ") AND ((pro_ordenproddet.estado)>=2));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    MENSAJE_ = "Solicitud de Materiales"
    
    If xRs.State = 0 Then verificarCambioEstado = False: GoTo SALIR_
    If xRs.RecordCount > 0 Then verificarCambioEstado = False: GoTo SALIR_
    
    ' Buscando Registros de Planilla
    cSQL = "SELECT pro_controltardet.idprocorr, pro_controltardet.estado " _
        + vbCr + "FROM pro_controltardet " _
        + vbCr + "WHERE (((pro_controltardet.idprocorr)=" & IDPROCORR_ & ") AND ((pro_controltardet.estado)>=2));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    MENSAJE_ = "Registros de Planilla"
    
    If xRs.State = 0 Then verificarCambioEstado = False: GoTo SALIR_
    If xRs.RecordCount > 0 Then verificarCambioEstado = False: GoTo SALIR_
    
    ' Buscando Registros de Almacen
    cSQL = "SELECT alm_ingreso.idprocorr, alm_ingreso.estado " _
        + vbCr + "FROM alm_ingreso " _
        + vbCr + "WHERE (((alm_ingreso.idprocorr)=" & IDPROCORR_ & ") AND ((alm_ingreso.estado)>=2));"
    
    Set xRs = Nothing
    RST_Busq xRs, cSQL, xCon
    
    MENSAJE_ = "Registros de Almacen"
    
    If xRs.State = 0 Then verificarCambioEstado = False: GoTo SALIR_
    If xRs.RecordCount > 0 Then verificarCambioEstado = False: GoTo SALIR_
    
    verificarCambioEstado = True
    Exit Function
    
SALIR_:
    MENSAJE_ = "Se han encontrado " & MENSAJE_ & " que se encuentran en un estado no modificable; " _
    & vbCr & "verifique la condición de dichos Registros para completar esta acción."
End Function

Sub Bloquea()
    Dim IDCRDET_ As Double
    Dim DISEÑO_ As Boolean
    
    DISEÑO_ = Not CalCtrlCronog.Visible
    
    TxtIdSup.Locked = Not TxtIdSup.Locked
    TxtFchIni.Locked = Not TxtFchIni.Locked
    TxtFchFin.Locked = Not TxtFchFin.Locked
    CmdBusSup.Enabled = Not CmdBusSup.Enabled
    habilitar Cmd, Not Cmd(0).Enabled
    
    Cmd(3).Enabled = True
    Cmd(11).Enabled = True
    
    '**************************************
    ' Se (In)habilita Estado
    Frame1.Enabled = Not Frame1.Enabled
    '**************************************
    
    ' Boton Imprimir
    If QueHace = 3 Then
        Cmd(19).Enabled = True
    Else
        Cmd(19).Enabled = False
    End If
    ' Boton Procesar
    CmdOpciones(0).Enabled = False
    
    If DISEÑO_ Then
        If fg(3).Row < fg(3).FixedRows Then Exit Sub
        IDCRDET_ = NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_))
        CmdOpciones(2).Enabled = False
        SliderCal.Enabled = False
    Else
        IDCRDET_ = NulosN(LblIdCrDet.Caption)
        CmdOpciones(2).Enabled = True
        SliderCal.Enabled = True
    End If
    
    If frm(2).Visible Then
        bloquearControles IDCRDET_, DISEÑO_
    End If
End Sub

Private Sub bloquearControles(IDCRDET_ As Double, Optional DISEÑO_ As Boolean = False)
    Dim HABILITADO_ As Boolean
    
    ' Se verifica el estado para bloquear
    If QueHace = 3 Then
        HABILITADO_ = False
    Else
        If RstProductos.State = 0 Then Exit Sub
        
        RstProductos.Filter = "id = " & IDCRDET_
        If RstProductos.RecordCount = 0 Then
            HABILITADO_ = True
        Else
            If NulosN(RstProductos("estado")) = ESTADOPENDIENTE_ Then
                HABILITADO_ = True
            Else
                HABILITADO_ = False
            End If
        End If
    End If
    
    If DISEÑO_ Then
        Frame6.Top = 350
        Frame6.ZOrder (0)
        
        TabOne2.Top = 3000
        TabOne2.Height = 4500
        
'        Frame8.Top = 2900
'        Frame8.Height = 4060
        fg(1).Height = 4050
        fg(4).Height = 4050
        Label17.Top = 4100
        LblDetTrab.Top = 4040
    Else
        Cmd(0).Enabled = HABILITADO_ ' Producto
        TxtMatProd.Locked = Not HABILITADO_
        Cmd(16).Enabled = HABILITADO_ ' Receta
        TxtCodRec.Locked = Not HABILITADO_
        Cmd(18).Enabled = HABILITADO_ ' Encargado
        TxtIdEncarg.Locked = Not HABILITADO_
        Cmd(20).Enabled = HABILITADO_ ' Linea
        TxtIdLineaDet.Locked = Not HABILITADO_
        TxtCant.Locked = Not HABILITADO_ ' Cantidad
        DTPHoras.Enabled = HABILITADO_ ' Hora de Inicio
        
        Frame6.Top = 1500
        TabOne2.Top = 4150
        TabOne2.Height = 3300
        fg(1).Height = 2835
        fg(4).Height = 2835
        Label17.Top = 2850
        LblDetTrab.Top = 2790
    End If
    
    Cmd(1).Enabled = HABILITADO_ ' Propiedades de Procesado
    Cmd(2).Enabled = HABILITADO_ ' Procesar
    Cmd(4).Enabled = HABILITADO_ ' Agregar
    Cmd(5).Enabled = HABILITADO_ ' Seleccionar
    Cmd(8).Enabled = HABILITADO_ ' Ranking
    Cmd(3).Enabled = HABILITADO_ ' Grupo
    Cmd(6).Enabled = HABILITADO_ ' Eliminar
    Cmd(7).Enabled = HABILITADO_ ' Eliminar Todos
    'cmd(10).Enabled = HABILITADO_ ' Aceptar
    
End Sub

Sub Blanquea()
    TxtIdSup.Text = ""
    LblSupervisor.Caption = ""
    TxtFchIni.valor = Date
    CalCtrlCronog.DataProvider.RemoveAllEvents
    If QueHace = 1 Then
        CalCtrlCronog.Visible = False
        cbfecha.Visible = False
        fg(3).Visible = False
        CmdOpciones(5).Enabled = False ' Boton cambio de Vista
        CmdOpciones(0).Enabled = True ' Boton Procesar
    End If
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
    
    Frame2.Width = TabOne1.Width - 1470
    LblSupervisor.Width = TabOne1.Width - 3615
    CmdOpciones(0).Left = TabOne1.Width - 1425
        
    fg(3).Width = TabOne1.Width - 100
    fg(3).Height = TabOne1.Height - 5205
    
    Frame6.Width = TabOne1.Width - 135
    Frame6.Top = TabOne1.Height - 3585
    
    fg(0).Width = Frame6.Width - 1605
    Cmd(1).Left = Frame6.Width - 1485
    Cmd(2).Left = Frame6.Width - 1485

    FrmBotones.Top = TabOne1.Height - 1000
    FrmBotones.Width = TabOne1.Width - 100
    
    SliderCal.Left = FrmBotones.Width - 2535
End Sub

Private Sub llenarComboFechas()
    Dim FECHAINI_ As Date
    Dim FECHAFIN_ As Date
    Dim FECHAAUX_ As Date
    Dim A As Integer
    
    FECHAINI_ = CDate(TxtFchIni.valor)
    FECHAFIN_ = CDate(TxtFchFin.valor)
    FECHAAUX_ = FECHAINI_
    
    cbfecha.Clear
    cbfecha.AddItem "TODOS"
    cbfecha.ItemData(cbfecha.NewIndex) = 0
    For A = 0 To (FECHAFIN_ - FECHAINI_)
        cbfecha.AddItem Format(FECHAAUX_, FORMAT_DATE)
        cbfecha.ItemData(cbfecha.NewIndex) = A + 1
        FECHAAUX_ = FECHAAUX_ + 1
    Next
    cbfecha.ListIndex = 0
End Sub

Sub Cancelar()
    QueHace = 3
    Bloquea
    Label5.Caption = "Consultando Cronograma de Produccion"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    
    CalCtrlCronog.Visible = True
    fg(3).Visible = True
    cbfecha.Visible = True
    ComboSemanas.Locked = True
    
    CmdOpciones(0).Enabled = False
    CmdOpciones(1).Enabled = False
    CmdOpciones(2).Enabled = False
    CmdOpciones(3).Enabled = False
    CmdOpciones(5).Enabled = True
    
    ActivaTool
End Sub

Private Sub operaciones(Optional AGREGAR_ As Boolean = True, Optional MODIFICAR_ As Boolean = False, _
                                                            Optional ELIMINAR_ As Boolean = False, _
                                                            Optional IDCRONODET_ As Double = 0)
    Dim IDCRDET_ As Double
    Dim Rpta As Integer
    
    IDCRDET_ = IDCRONODET_
    If AGREGAR_ Then
        CORR_ = CORR_ + 1
        llenarDatos False, IDCRDET_
        frm(2).Visible = False
    End If
    
    If MODIFICAR_ Then
    End If
    
    If ELIMINAR_ Then
        
        If DETECTOR_.ViewEvent Is Nothing Then Exit Sub
        ' Se encuentra el idcrdet involucrado
        IDCRDET_ = EVENTO_.ReminderSoundFile
        ' Se verifica el estado del recordset
        If RstProductos.State = 0 Then Exit Sub
        RstProductos.Filter = "id = " & IDCRDET_ & ""
        If RstProductos.RecordCount = 0 Then Exit Sub
                
        If NulosN(RstProductos("cerrado")) = -1 Then
            MsgBox "No se puede eliminar un registro aprobado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
                
        ' Se llena el recordset auxiliar
        If RstProductosAux.State = 0 Then DEFINIR_RST_TMP RstProductosAux, RstProductos
        limpiarRST RstProductosAux
        CARGAR_RST_TMP RstProductosAux, RstProductos
        limpiarRST RstProductos, False
        
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
            CalCtrlCronog.DataProvider.DeleteEvent EVENTO_
        Else
            CARGAR_RST_TMP RstProductos, RstProductosAux
            RstProductos.Filter = adFilterNone
        End If
            
        ' Se limpia el calendario
        CalCtrlCronog.DataProvider.RemoveAllEvents
        ' Se llenan todos los eventos
        llenarDatos
    End If
    
    Set DETECTOR_ = Nothing
End Sub

Private Function HallaValor(conn As ADODB.Connection, tabla As String, campo As String) As Long
    Dim xRs As New ADODB.Recordset
    On Error GoTo error
    RST_Busq xRs, "SELECT top 1 CLng([" + campo + "]) AS num FROM " + tabla + " ORDER BY CLng([" + campo + "]) DESC;", conn
    If xRs.State = 1 Then
        If xRs.EOF = False And xRs.BOF = False And xRs.RecordCount <> 0 Then
            HallaValor = NulosN(xRs.Fields(0)) + 1
        End If
    Else
        HallaValor = -1
    End If
    Set xRs = Nothing
    Exit Function

error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "HallarValor"
End Function

Function GrabarProduccion(IDCRDET_ As Double, ByRef IDREGPROD_ As Double) As Boolean
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstIns As New ADODB.Recordset
    Dim RstTar As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim xId As Double
    Dim xCodDet&
    Dim xCol&, xFil&
    Dim xCorr&
    Dim xRs As New ADODB.Recordset
    
On Error GoTo LaCague
    If RstProductos.State = 0 Then GrabarProduccion = False: Exit Function
    RstProductos.Filter = "id = " & IDCRDET_
    If RstProductos.RecordCount = 0 Then GrabarProduccion = False: Exit Function

    'xCon.BeginTrans
    Me.MousePointer = vbHourglass
        
    cSQL = "SELECT pro_produccion.id, pro_produccion.dia, pro_produccion.idsup, pro_producciondet.idres " _
        + vbCr + "FROM pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro " _
        + vbCr + "WHERE (((pro_produccion.dia)=CDate('" & Format(RstProductos("fchpro"), "dd/mm/yyyy") & "')) AND ((pro_produccion.idsup)=" & NulosN(TxtIdSup.Text) & ") AND ((pro_producciondet.idres)=" & NulosN(RstProductos("idresp")) & "));"
    
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then
        GrabarProduccion = False
        Me.MousePointer = vbDefault
        Exit Function
    End If
    
    Dim IDPRODETCORR_ As Double
    IDPRODETCORR_ = HallaCodigoTabla("pro_producciondet", xCon, "corr")
    
    If xRs.RecordCount = 0 Then
        RST_Busq RstCab, "SELECT top 1 * FROM pro_produccion ", xCon
        xId = HallaCodigoTabla("pro_produccion", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = xRs("id")
        RST_Busq RstCab, "SELECT * FROM pro_produccion WHERE id = " & xId & ";", xCon

        ' restar el stock actual encabezado
        cSQL = "SELECT pro_producciondet.iditem, Sum(pro_producciondet.cantidad) AS total " _
        + vbCr + "FROM pro_producciondet AS pro_producciondet " _
        + vbCr + "GROUP BY pro_producciondet.idpro, pro_producciondet.iditem, pro_producciondet.idcrdet " _
        + vbCr + "HAVING (((pro_producciondet.idpro)=" & xId & ") And ((pro_producciondet.idcrdet)= " & IDCRDET_ & "));"
        
        RST_Busq RstTmp, cSQL, xCon
        If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            xCon.Execute "UPDATE alm_inventario SET alm_inventario.stckact = [alm_inventario].[stckact] - " & NulosN(RstTmp("total")) & " WHERE (((alm_inventario.id)=" & RstTmp("iditem") & "));"
            RstTmp.MoveNext
        Loop
        Set RstTmp = Nothing

        ' acumular el stock actual detalle
        cSQL = "SELECT pro_producciondetins.iditem, Sum(pro_producciondetins.canutil) AS total " _
        + vbCr + "FROM pro_producciondet INNER JOIN pro_producciondetins ON (pro_producciondet.idrec = pro_producciondetins.idrec) AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro) " _
        + vbCr + "GROUP BY pro_producciondetins.iditem, pro_producciondet.idpro, pro_producciondet.idcrdet " _
        + vbCr + "HAVING (((pro_producciondet.idpro)=" & xId & ") And ((pro_producciondet.idcrdet)= " & IDCRDET_ & "));"
        
        RST_Busq RstTmp, cSQL, xCon
        If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            xCon.Execute "UPDATE alm_inventario SET alm_inventario.stckact = [alm_inventario].[stckact] + " & NulosN(RstTmp("total")) & " WHERE (((alm_inventario.id)=" & RstTmp("iditem") & "));"
            RstTmp.MoveNext
        Loop
        Set RstTmp = Nothing

        ' eliminando los registros
        xCon.Execute "DELETE * FROM pro_producciondetins WHERE ((idpro = " & xId & ") And (idcrdet = " & IDCRDET_ & "))"
        xCon.Execute "DELETE * FROM pro_producciondettar WHERE ((idpro = " & xId & ") And (idcrdet = " & IDCRDET_ & "))"
        xCon.Execute "DELETE * FROM pro_producciondet WHERE ((idpro = " & xId & ") And (idcrdet = " & IDCRDET_ & "))"

        '--actualizando el codigo de produccion del programa de produccion a 0
        xCon.Execute "UPDATE pro_programadet SET idpro =0 WHERE idpro = " & xId & ""
    End If
    
    mIdRegistro = xId
    
    RST_Busq RstDet, "SELECT top 1 * FROM pro_producciondet", xCon
    RST_Busq RstIns, "SELECT top 1 * FROM pro_producciondetins", xCon
    RST_Busq RstTar, "SELECT top 1 * FROM pro_producciondettar", xCon
    
    RstCab("dia") = Format(RstProductos("fchpro"), "dd/mm/yyyy")
    RstCab("idsup") = NulosN(TxtIdSup.Text) 'NulosN(RstProductos("idresp"))
    RstCab("num") = Format(xId, "000000")
    RstCab("obs") = ""
    
    RstCab.Update
    
    Dim F_CAMBIO_PRODUCCION As Boolean
    Dim M_PRODUCCION As String
    Dim CODIGORESP_ As Double
    
    CODIGORESP_ = Busca_Codigo(NulosN(RstProductos("idresp")), "idemp", "id", "pro_emp", "N", xCon)
    
    RstDet.AddNew
    
    '*****************************
    RstDet("corr") = IDPRODETCORR_
    '*****************************
    
    RstDet("idpro") = xId
    M_PRODUCCION = Format(HallaValor(xCon, "pro_producciondet", "numparte"), "00000000")
    RstDet("numparte") = M_PRODUCCION
    RstDet("idrec") = NulosN(RstProductos("idrec"))
    RstDet("iditem") = NulosN(RstProductos("iditem"))
    RstDet("idunimed") = Busca_Codigo(NulosN(RstProductos("iditem")), "id", "idunimed", "alm_inventario", "N", xCon) 'NulosN(RstProductos("abrev"))
    RstDet("cantidad") = NulosN(RstProductos("cantidad"))
    RstDet("horini") = RstProductos("horpro")
    RstDet("horfin") = RstProductos("horfin")
    RstDet("idres") = CODIGORESP_
    RstDet("idturno") = 1
    RstDet("canprog") = NulosN(RstProductos("cantidad"))
    RstDet("numprog") = NulosN(RstProductos("numprod"))
    RstDet("idcrdet") = IDCRDET_
    
    '*************************
    IDREGPROD_ = IDPRODETCORR_
    RstDet("estado") = 1
    '*************************
    
    RstDet("obs") = ""
    RstDet.Update
    
    Dim RST_INSUMO As New ADODB.Recordset
        
    cSQL = "SELECT pro_recetains.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro, [pro_recetains]![canpro]* " & NulosN(RstProductos("cantidad")) & " AS canreq, pro_recetains.idunimed " _
        + vbCr + " FROM (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
        + vbCr + " WHERE (((pro_recetains.idrec)=" & NulosN(RstProductos("idrec")) & "));"
    
    RST_Busq RST_INSUMO, cSQL, xCon
    
    If RST_INSUMO.RecordCount = 0 Then
        GrabarProduccion = False
        Me.MousePointer = vbDefault
        Exit Function
    End If
    RST_INSUMO.MoveFirst
    
    Do While Not RST_INSUMO.EOF
        RstIns.AddNew
        RstIns("idpro") = xId
        RstIns("numparte") = M_PRODUCCION
        RstIns("idrec") = NulosN(RstProductos("idrec"))
        RstIns("iditem") = NulosN(RST_INSUMO.Fields("iditem"))
        RstIns("idunimed") = NulosN(RST_INSUMO.Fields("idunimed"))
        RstIns("canutil") = NulosN(RST_INSUMO.Fields("canreq"))
        RstIns("canpro") = NulosN(RST_INSUMO.Fields("canpro"))
        RstIns("idcrdet") = IDCRDET_
        RstIns.Update
        
        RST_INSUMO.MoveNext
    Loop
    
    RstTareas.Filter = "idcrdet = " & IDCRDET_ & " And activo = -1"
    If RstTareas.RecordCount = 0 Then
        GrabarProduccion = False
        Me.MousePointer = vbDefault
        Exit Function
    End If
    
    RstTareas.MoveFirst
    xCorr = 1
    Do While Not RstTareas.EOF
        RstTar.AddNew
        '--CLAVE
        RstTar("idpro") = xId
        RstTar("numparte") = M_PRODUCCION
        RstTar("idrec") = NulosN(RstProductos("idrec"))
        RstTar("idtar") = NulosN(RstTareas("idtar"))
        RstTar("corr") = xCorr
        RstTar("idunimed") = 2
        RstTar("horini") = CDate(RstTareas.Fields("horinitar"))
        RstTar("horfin") = CDate(RstTareas.Fields("horfintar"))
        RstTar("canper") = NulosN(RstTareas.Fields("numper"))
        RstTar("idcrdet") = IDCRDET_
        RstTar.Update
        RstTareas.MoveNext
        xCorr = xCorr + 1
    Loop
    
    'xCon.CommitTrans
    GrabarProduccion = True

SALIR:
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstIns = Nothing:    Set RstTar = Nothing:    Set RstTmp = Nothing
    Exit Function

LaCague:
    'xCon.RollbackTrans
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstIns = Nothing:    Set RstTar = Nothing:    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
    GrabarProduccion = False
End Function

Function GrabarAlmacen(IDCRDET_ As Double, IDREGSOL_ As Double, IDREGPROD_ As Double) As Boolean
    Dim xId As Double
    Dim A As Integer
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim NUMERODOC_ As String
    Dim RstValores As New ADODB.Recordset
    Dim xRsAux As New ADODB.Recordset
    
On Error GoTo LaCague
    
    ' Se buscan si hay registros para grabar
    RstProductos.Filter = "id = " & IDCRDET_
    If RstProductos.RecordCount = 0 Then GrabarAlmacen = False: Exit Function
    
    ' Se halla el correlativo de Documento
    cSQL = "SELECT Max(alm_ingreso.numdoc) AS maxnum " _
        + vbCr + "From alm_ingreso " _
        + vbCr + "GROUP BY alm_ingreso.tipdoc " _
        + vbCr + "HAVING (((alm_ingreso.tipdoc)=71));"
    
    RST_Busq xRsAux, cSQL, xCon
    
    If xRsAux.State = 0 Then GrabarAlmacen = False: Exit Function
    If xRsAux.RecordCount = 0 Then
        NUMERODOC_ = 1
    Else
        NUMERODOC_ = NulosN(xRsAux("maxnum")) + 1
    End If
    
    ' Se buscan si hay registros para la produccion Actual
    cSQL = "SELECT alm_ingreso.id, alm_ingreso.idprocorr " _
        + vbCr + "FROM alm_ingreso " _
        + vbCr + "WHERE (((alm_ingreso.idprocorr)=" & IDREGPROD_ & "));"
    
    RST_Busq xRs, cSQL, xCon
    If xRs.State = 0 Then GrabarAlmacen = False: Exit Function
    
    If xRs.RecordCount = 0 Then
        xId = HallaCodigoTabla("alm_ingreso", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM alm_ingreso", xCon
        RST_Busq RstDet, "SELECT TOP 1 * FROM alm_ingresodet", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xRs.MoveFirst
        While Not xRs.EOF
            xId = NulosN(xRs("id"))
            RST_Busq RstCab, "SELECT * FROM alm_ingreso WHERE id = " & xId, xCon
            xCon.Execute "DELETE * FROM alm_ingresodet WHERE (id = " & xId & ")"
            RST_Busq RstDet, "SELECT * FROM alm_ingresodet", xCon
            
            xRs.MoveNext
        Wend
    End If
    
    ' Se graba la cabecera
    RstCab("tipdoc") = 71
    RstCab("fching") = Format(RstProductos("fchpro"), "dd/mm/yyyy")
    RstCab("fchdoc") = Format(RstProductos("fchpro"), "dd/mm/yyyy")
    RstCab("numser") = "0001"
    RstCab("numdoc") = Format(NUMERODOC_, "0000000000")
    RstCab("idres") = NulosN(RstProductos("idresp"))
    RstCab("nombre") = NulosC(RstProductos("descripcion"))
    RstCab("tipmov") = 0
    RstCab("ano") = AnoTra
    RstCab("idmes") = Month(RstProductos("fchpro"))
    RstCab("idorddet") = IDREGSOL_
    RstCab("estado") = 1
    '*********************************
    RstCab("idtipdocref") = 110
    RstCab("iddocref") = IDREGSOL_
    '*********************************
    RstCab.Update
        
    ' Se graba el detalle
    cSQL = "SELECT pro_recetains.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro, [pro_recetains]![canpro]* " & NulosN(RstProductos("cantidad")) & " AS canreq, alm_inventario.tippro " _
        + vbCr + " FROM (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
        + vbCr + " WHERE (((pro_recetains.idrec)=" & NulosN(RstProductos("idrec")) & "));"
    
    RST_Busq RstValores, cSQL, xCon
    If RstValores.RecordCount = 0 Then GrabarAlmacen = False: Exit Function
    RstValores.MoveFirst
    For A = 1 To RstValores.RecordCount
        RstDet.AddNew
        RstDet("id") = xId
        RstDet("iditem") = NulosN(RstValores("iditem"))
        RstDet("cantidad") = 0
        RstDet("cantteo") = NulosN(RstValores("canreq"))
        RstDet("idcrdet") = IDCRDET_
        RstDet("idtipo") = NulosN(RstValores("tippro"))
        RstDet.Update
        
        RstValores.MoveNext
    Next A
    
    '*************************************************************************
    ' Se crea la salida para el reproceso
    If RstReproc.State = 0 Then GrabarAlmacen = False: GoTo SALIR_
    RstReproc.Filter = "idcrdet = " & IDCRDET_
    If RstReproc.RecordCount = 0 Then GrabarAlmacen = False: GoTo SALIR_
    RstReproc.MoveFirst
    While Not RstReproc.EOF
        xId = xId + 1
        NUMERODOC_ = NUMERODOC_ + 1
        IDREGSOL_ = IDREGSOL_ + 1
        RstCab.AddNew
        RstCab("id") = xId
        RstCab("tipdoc") = 71
        RstCab("fching") = Format(RstProductos("fchpro"), "dd/mm/yyyy")
        RstCab("fchdoc") = Format(RstProductos("fchpro"), "dd/mm/yyyy")
        RstCab("numser") = "0001"
        RstCab("numdoc") = Format(NUMERODOC_, "0000000000")
        RstCab("idres") = NulosN(RstProductos("idresp"))
        RstCab("nombre") = NulosC(RstProductos("descripcion")) & " / REPROCESO"
        RstCab("tipmov") = 0
        RstCab("ano") = AnoTra
        RstCab("idmes") = Month(RstProductos("fchpro"))
        RstCab("idorddet") = IDREGSOL_
        RstCab("estado") = 1
        RstCab.Update
        
        ' se Graba el detalle
        RstDet.AddNew
        RstDet("id") = xId
        RstDet("iditem") = NulosN(RstProductos("iditem"))
        RstDet("cantidad") = 0
        RstDet("cantteo") = NulosN(RstReproc("cantidad"))
        RstDet("idtipo") = Busca_Codigo(NulosN(RstProductos("iditem")), "id", "tippro", "alm_inventario", "N", xCon)
        RstDet.Update
        
        RstReproc.MoveNext
    Wend
    '*************************************************************************
    
SALIR_:
    GrabarAlmacen = True
    Exit Function

LaCague:
    'Resume
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    GrabarAlmacen = False
    Exit Function
End Function

Function grabarSolicitud(IDCRDET_ As Double, ByRef IDREGSOL_ As Double, IDREGPROD_ As Double) As Boolean
    Dim A As Integer
    Dim B As Integer
    Dim xTot As Long
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstDetIns As New ADODB.Recordset
    Dim xId As Double
    Dim nSQL As String
    Dim xIdDet As Integer
    Dim xRs As New ADODB.Recordset
    Dim RstValores As New ADODB.Recordset
    Dim NUMDOC_ As Integer
    
    If RstProductos.State = 0 Then grabarSolicitud = False: Exit Function
    RstProductos.Filter = "id = " & IDCRDET_
    If RstProductos.RecordCount = 0 Then grabarSolicitud = False: Exit Function
   
On Error GoTo LaCague

    Me.MousePointer = vbHourglass
    
    Dim FECHA_ As String
    Dim IDSUP_ As Integer
    Dim IDRESP_ As Integer
    
    FECHA_ = Format(RstProductos("fchpro"), "dd/mm/yyyy")
    IDSUP_ = NulosN(TxtIdSup.Text)
    IDRESP_ = NulosN(RstProductos("idresp"))
    
    ' Se buscan registros semejantes
    cSQL = "SELECT pro_ordenprod.id, pro_ordenprod.fchemi, pro_ordenprod.idsup, pro_ordenproddet.idresp " _
        + vbCr + "FROM pro_ordenprod LEFT JOIN pro_ordenproddet ON pro_ordenprod.id = pro_ordenproddet.idord " _
        + vbCr + "WHERE (((pro_ordenprod.fchemi)=CDate('" & FECHA_ & "')) " _
                            & "AND ((pro_ordenprod.idsup)=" & IDSUP_ & ") " _
                            & "AND ((pro_ordenproddet.idresp)=" & IDRESP_ & "));"
    
    RST_Busq xRs, cSQL, xCon
    If xRs.State = 0 Then grabarSolicitud = False: Exit Function
    
    If xRs.RecordCount = 0 Then
        xId = HallaCodigoTabla("pro_ordenprod", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM pro_ordenprod", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = xRs("id")
        RST_Busq RstCab, "SELECT * FROM pro_ordenprod WHERE id = " & xId & "", xCon
    End If
    
    RST_Busq RstDet, "SELECT TOP 1 * FROM pro_ordenproddet", xCon
    RST_Busq RstDetIns, "SELECT TOP 1 * FROM pro_ordenproddetins", xCon
    
    xIdDet = HallaCodigoTabla("pro_ordenproddet", xCon, "id")
    NUMDOC_ = HallaCodigoTabla("pro_ordenproddet", xCon, "numdoc")
    
    ' Se llena cabecera
    RstCab("idsup") = NulosN(TxtIdSup.Text)
    RstCab("fchemi") = Format(RstProductos("fchpro"), "dd/mm/yyyy")
    RstCab.Update
    ' Se llena el detalle
    RstDet.AddNew
    RstDet("id") = xIdDet
    RstDet("idord") = xId
    RstDet("iditem") = NulosN(RstProductos("iditem"))
    RstDet("idrec") = NulosN(RstProductos("idrec"))
    RstDet("idunimed") = Busca_Codigo(NulosN(RstProductos("iditem")), "id", "idunimed", "alm_inventario", "N", xCon) 'NulosN(RstProductos("abrev"))
    RstDet("cantidad") = NulosN(RstProductos("cantidad"))
    RstDet("numdoc") = Format(NUMDOC_, "000000")
    RstDet("fchprog") = Format(RstProductos("fchpro"), "dd/mm/yyyy")
    RstDet("idresp") = NulosN(RstProductos("idresp"))
    RstDet("tipo") = 1
    RstDet("idprocorr") = IDREGPROD_
    RstDet("estado") = 1
    RstDet("obs") = ""
    RstDet.Update
    
    ' Se llenan los Insumos
    IDREGSOL_ = xIdDet
    cSQL = "SELECT pro_recetains.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro, [pro_recetains]![canpro]* " & NulosN(RstProductos("cantidad")) & " AS canreq " _
        + vbCr + " FROM (pro_recetains LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
        + vbCr + " WHERE (((pro_recetains.idrec)=" & NulosN(RstProductos("idrec")) & "));"
    
    RST_Busq RstValores, cSQL, xCon
    If RstValores.RecordCount = 0 Then grabarSolicitud = False: Exit Function
    RstValores.MoveFirst
    For B = 1 To RstValores.RecordCount
        RstDetIns.AddNew
        RstDetIns("idord") = xId
        RstDetIns("idorddet") = xIdDet
        RstDetIns("activo") = -1
        RstDetIns("iditem") = NulosN(RstValores("iditem"))
        RstDetIns("cantidad") = NulosN(RstValores("canreq"))
        RstDetIns.Update
        RstValores.MoveNext
        If RstValores.EOF Then Exit For
    Next B
    
    ' Se llenan los Reprocesos
    If RstReproc.State = 0 Then GoTo SALIR_
    RstReproc.Filter = "idcrdet=" & IDCRDET_
    If RstReproc.RecordCount = 0 Then GoTo SALIR_
    xIdDet = xIdDet + 1
    NUMDOC_ = NUMDOC_ + 1
    ' Se llena el detalle
    RstDet.AddNew
    RstDet("id") = xIdDet
    RstDet("idord") = xId
    RstDet("numdoc") = Format(NUMDOC_, "000000")
    RstDet("fchprog") = Format(RstProductos("fchpro"), "dd/mm/yyyy")
    RstDet("idresp") = RstProductos("idresp")
    RstDet("tipo") = 2
    RstDet("idprocorr") = IDREGPROD_
    RstDet("estado") = 2
    RstDet("obs") = "REPROCESO"
    RstDet.Update
    
    ' Se llenan los Insumos
    RstReproc.MoveFirst
    While Not RstReproc.EOF
        RstDetIns.AddNew
        RstDetIns("idord") = xId
        RstDetIns("idorddet") = xIdDet
        RstDetIns("activo") = -1
        RstDetIns("iditem") = NulosN(RstProductos("iditem"))
        RstDetIns("cantidad") = NulosN(RstReproc("cantidad"))
        RstDetIns("idlotedet") = NulosN(RstReproc("idlotedet"))
        RstDetIns.Update
        
        RstReproc.MoveNext
    Wend
    
SALIR_:
    Me.MousePointer = vbDefault
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDetIns = Nothing
    grabarSolicitud = True
    
    Exit Function
LaCague:
    Me.MousePointer = vbDefault
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstDetIns = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    grabarSolicitud = False
End Function

Private Function definirUnirPersonal(ByRef RSTORIGEN_ As ADODB.Recordset, ByRef RSTTIPO_ As ADODB.Recordset) As ADODB.Recordset
    Dim xRs As New ADODB.Recordset
        
    DEFINIR_RST_TMP xRs, RSTTIPO_
    If RSTORIGEN_.State = 0 Then Exit Function
    If RSTORIGEN_.RecordCount = 0 Then Exit Function
    
    RSTORIGEN_.MoveFirst
    RSTTIPO_.Filter = adFilterNone
    While Not RSTORIGEN_.EOF
        RSTTIPO_.Filter = "idcrdet=" & (RSTORIGEN_("idcrdet")) & " And idtar=" & NulosN(RSTORIGEN_("idtar"))
        If RSTTIPO_.RecordCount = 0 Then GoTo SIGUIENTE_
        RSTTIPO_.MoveFirst
        While Not RSTTIPO_.EOF
            xRs.AddNew
            xRs("idper") = NulosN(RSTTIPO_("idper"))
            xRs("idtar") = NulosN(RSTTIPO_("idtar"))
            xRs("nombre") = NulosC(RSTTIPO_("nombre"))
            xRs.Update
            
            RSTTIPO_.MoveNext
        Wend
SIGUIENTE_:
        RSTORIGEN_.MoveNext
    Wend
    
    Set definirUnirPersonal = xRs
End Function

Function GrabarPlanilla(IDCRDET_ As Double, IDREGPROD_ As Double) As Boolean
    Dim xRs As New ADODB.Recordset
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstGr As New ADODB.Recordset
    Dim RstDetTara As New ADODB.Recordset
    Dim RstGrTara As New ADODB.Recordset
    Dim xId As Double
    Dim xCol&, xFil&, xItem&
    Dim HoraFraccion As Double
    Dim Difhora As String
    Dim RstDetTareas As New ADODB.Recordset
    Dim CODIGORESP_ As Double
    Dim CODIGOAREA_ As Double
    Dim CONTADOR_ As Integer
    Dim RstAux As New ADODB.Recordset
    Dim IDRESP_ As Double
    Dim IDAREA_ As Double
    Dim RstTarAux As New ADODB.Recordset
    Dim HORINI_ As String
    Dim HORFIN_ As String
    Dim CANT_ As Double
    Dim CAMBIO_ As Boolean
    Dim A As Integer
    Dim xRsTarAuxAux As New ADODB.Recordset
    Dim FCHPRO_ As String
    Dim NUEVO_ As Boolean

On Error GoTo LaCague

    RstProductos.Filter = "id = " & IDCRDET_
    If RstProductos.State = 0 Then Me.MousePointer = vbDefault: GrabarPlanilla = False: Exit Function
    If RstProductos.RecordCount = 0 Then Me.MousePointer = vbDefault: GrabarPlanilla = False: Exit Function
            
    ' Se filtra las tareas involucradas
    RstTareas.Filter = "idcrdet = " & IDCRDET_ & " And activo=-1"
    
    If (RstTareas.State = 0 Or RstPersonal.State = 0) Then Me.MousePointer = vbDefault: GrabarPlanilla = False: Exit Function
    If (RstTareas.RecordCount = 0 Or RstPersonal.RecordCount = 0) Then Me.MousePointer = vbDefault: GrabarPlanilla = False: Exit Function
    
    mIdRegistro = xId
    RST_Busq RstDet, "SELECT top 1 * FROM pro_controltardet", xCon
    RST_Busq RstDetTara, "SELECT top 1 * FROM pro_controltardetpes", xCon
    RST_Busq RstGr, "SELECT top 1 * FROM pro_controltardetgr", xCon
    RST_Busq RstGrTara, "SELECT top 1 * FROM pro_controltardetgrpes", xCon
    RST_Busq RstDetTareas, "SELECT top 1 * FROM pro_controltardettar", xCon
    
    ' Se define el recordset de tareas auxiliar
    DEFINIR_RST_TMP RstTarAux, RstTareas
    CARGAR_RST_TMP RstTarAux, RstTareas
        
    RstTareas.MoveFirst
    IDAREA_ = NulosN(RstTareas("idarea"))
    IDRESP_ = NulosN(RstTareas("idresp"))
    CODIGORESP_ = Busca_Codigo(IDRESP_, "idemp", "id", "pro_emp", "N", xCon)
    CODIGOAREA_ = IDAREA_
    FCHPRO_ = Format(RstTareas("fchini"), "dd/mm/yyyy")
    CAMBIO_ = True
    NUEVO_ = True
    
    Me.MousePointer = vbHourglass
    While Not RstTareas.EOF
        If Not CAMBIO_ Then GoTo SIGUIENTE_
        
        '**********************************************************************************
        cSQL = "SELECT pro_controltar.* " _
            + vbCr + "FROM pro_controltar " _
            + vbCr + "WHERE (fchtra = CDate('" & FCHPRO_ & "') And tipo = 2 " _
                                    & "And idres = " & CODIGORESP_ & " And estado=" & ESTADOPENDIENTE_ & ")"
        
        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 0 Then Me.MousePointer = vbDefault: GrabarPlanilla = False: Exit Function
        
        If xRs.RecordCount = 0 Then
            If NUEVO_ Then
                cSQL = "SELECT TOP 1 * FROM pro_controltar"
                RST_Busq RstCab, cSQL, xCon
                xId = HallaCodigoTabla("pro_controltar", xCon, "id")
            Else
                xId = xId + 1
            End If
            
            RstCab.AddNew
            RstCab("id") = xId
            RstCab("fchtra") = FCHPRO_
            RstCab("idarea") = CODIGOAREA_
            RstCab("idres") = CODIGORESP_
            RstCab("tipo") = 2
            RstCab("ano") = AnoTra
            RstCab("idmes") = Month(RstProductos("fchpro"))
            '*********************************
            RstCab("estado") = ESTADOPENDIENTE_
            '*********************************
            RstCab.Update
            
            CONTADOR_ = 0
        Else
            xId = xRs("id")
            If NUEVO_ Then
                cSQL = "SELECT * FROM pro_controltar WHERE id =" & xId & ""
                RST_Busq RstCab, cSQL, xCon
            End If
            
            cSQL = "SELECT Max(pro_controltardet.corr) AS maxcorr " _
                + vbCr + "FROM pro_controltardet " _
                + vbCr + "WHERE (((pro_controltardet.idctr)= " & xId & "))"
            
            RST_Busq RstAux, cSQL, xCon
            If RstAux.State = 0 Then Me.MousePointer = vbDefault: GrabarPlanilla = False: Exit Function
            If RstAux.RecordCount = 0 Then Me.MousePointer = vbDefault: GrabarPlanilla = False: Exit Function
            
            CONTADOR_ = NulosN(RstAux("maxcorr")) + 1
        End If
        '**********************************************************************************
        
        RstTarAux.Filter = "idarea=" & IDAREA_ & " And idresp=" & IDRESP_
        
        ' Se Obtiene la hora de inicio y de Fin
        RstTarAux.MoveFirst
        For A = 1 To RstTarAux.RecordCount
            ' Se llena hora de Inicio
            If A = 1 Then HORINI_ = Format(RstTarAux("horinitar"), "HH:mm")
            ' Se llena hora de fin y la cantidad
            If A = RstTarAux.RecordCount Then
                HORFIN_ = Format(RstTarAux("horfintar"), "HH:mm")
                CANT_ = NulosN(RstTarAux("cantproc"))
            End If
            RstTarAux.MoveNext
        Next A
        
        '**********************
        ' Detalle de Producto
        '**********************
        RstDet.AddNew
        RstDet("idctr") = xId
        RstDet("corr") = CONTADOR_
        RstDet("numlote") = ""
        RstDet("tipo") = 3
        RstDet("idref") = 0
        RstDet("idrec") = NulosN(RstProductos("idrec"))
        RstDet("idtar") = 0
        RstDet("horini") = CDate(HORINI_)
        RstDet("horfin") = CDate(HORFIN_)
        RstDet("cant") = CANT_
        RstDet("idunimed") = 2
        RstDet("observacion") = ""
        RstDet("observado") = 0
        RstDet("reproceso") = 0
        RstDet("cant1") = 0
        RstDet("idprocorr") = IDREGPROD_
        RstDet("estado") = 1
        RstDet.Update
        
        '**********************
        ' Detalle de Tareas
        '**********************
        RstTarAux.MoveFirst
        For A = 1 To RstTarAux.RecordCount
            RstDetTareas.AddNew
            RstDetTareas("idctr") = xId
            RstDetTareas("corr") = CONTADOR_
            RstDetTareas("idrec") = NulosN(RstProductos("idrec"))
            RstDetTareas("idtar") = NulosN(RstTarAux("idtar"))
            RstDetTareas("orden") = NulosN(RstTarAux("orden"))
            RstDetTareas("activo") = NulosN(RstTarAux("activo"))
            RstDetTareas("idcrdet") = IDCRDET_
            RstDetTareas.Update
            
            RstTarAux.MoveNext
        Next A
        
        '**********************
        ' Detalle de Personal
        '**********************
        ' Se une a todas la spersona involucradas
        DEFINIR_RST_TMP xRsTarAuxAux, RstTarAux
        CARGAR_RST_TMP xRsTarAuxAux, RstTarAux
        Set xRs = definirUnirPersonal(RstTarAux, RstPersonal)
        If xRs.RecordCount > 0 Then
            xRs.MoveFirst
            For A = 1 To xRs.RecordCount
                RstGr.AddNew
                RstGr("idctr") = xId
                RstGr("corr") = CONTADOR_
                RstGr("idper") = NulosN(xRs("idper"))
                RstGr("idrec") = NulosN(RstProductos("idrec"))
                RstGr("cant") = CANT_ / xRs.RecordCount
                RstGr("canpro") = CANT_ / xRs.RecordCount
                RstGr("cantbrut") = 0
                RstGr("activo") = True
                RstGr("horini") = CDate(HORINI_)
                RstGr("horfin") = CDate(HORFIN_)
                ' calculando las horas de trabajo
                'Difhora = DiferenciaHoras(HORINI_, HORFIN_, True)
                Difhora = Format(CDate(HORINI_) - CDate(HORFIN_), "HH:mm")
                HoraFraccion = Convert1HoraFaccion(Difhora)
                RstGr("tothor") = HoraFraccion
                If IsDate(Difhora) = True Then RstGr("difhor") = CDate(Difhora)
                RstGr("idcrdet") = IDCRDET_
                RstGr.Update
                
                xRs.MoveNext
            Next A
        End If
        
SIGUIENTE_:
        RstTareas.MoveNext
        
        ' se verifica si cambio el registro
        If Not RstTareas.EOF Then
            If IDAREA_ <> NulosN(RstTareas("idarea")) Or IDRESP_ <> NulosN(RstTareas("idresp")) Then
                CAMBIO_ = True
                CONTADOR_ = CONTADOR_ + 1
                IDAREA_ = NulosN(RstTareas("idarea"))
                IDRESP_ = NulosN(RstTareas("idresp"))
                CODIGORESP_ = Busca_Codigo(IDRESP_, "idemp", "id", "pro_emp", "N", xCon)
                CODIGOAREA_ = IDAREA_
                FCHPRO_ = Format(RstTareas("fchini"), "dd/mm/yyyy")
                
                CONTADOR_ = 0
                xId = xId + 1
            Else
                CAMBIO_ = False
            End If
        End If
    Wend
    
    GrabarPlanilla = True
SALIR:
    Me.MousePointer = vbDefault
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstGr = Nothing
    Set RstDetTareas = Nothing
    Exit Function

LaCague:
    'Resume
    Me.MousePointer = vbDefault
    Set RstCab = Nothing
    Set RstDet = Nothing
    Set RstGr = Nothing
    Set RstDetTareas = Nothing
    SHOW_ERROR Me.Name, "GrabarPlanilla", True, "No se pudo guardar el registro por el siguiente motivo :"
    GrabarPlanilla = False
End Function

Private Function hallarNumeroProduccion() As String
    Dim NUMERO_ As String
    NUMERO_ = 0
    
    hallarNumeroProduccion = NUMERO_
End Function

Private Sub mostrarFormulario(Optional AGREGAR_ As Boolean = True, Optional MODIFICAR_ As Boolean = False, _
                                                                    Optional RECETA_ As Boolean = False, _
                                                                    Optional DISEÑO_ As Boolean = False)
    Dim IDCRDET_ As Double
    Dim fIni As Date
    Dim fFin As Date
    Dim AllDay As Boolean
    Dim A As Integer
    
    If AGREGAR_ Then ' Nuevo
        If DISEÑO_ Then
            LblDia.Caption = Format(fg(3).TextMatrix(fg(3).Row, COLUMNAFCHPROD_), "dd/mm/yyyy")
            lblNumprod.Caption = fg(3).TextMatrix(fg(3).Row, COLUMNANUMPROD_)
        Else
            CalCtrlCronog.ActiveView.GetSelection fIni, fFin, AllDay
            
            ' SI es una fecha Incoherente
            If Format(fIni, "yyyy") < AnoTra Then Exit Sub
            
            LblDia.Caption = Format(fIni, "dd/mm/yyyy")         ' Detalle del Dia
            lblNumprod.Caption = hallarNumeroProduccion         ' Numero de Produccion
            LblIdCrDet.Caption = CORR_                          ' Correlativo
            TxtMatProd.Text = ""                                ' iditem
            LblMatProd.Caption = ""                             ' Descripcion prod
            TxtCodRec.Text = ""                                 ' codigo de receta
            LblIdRec.Caption = ""                               ' id receta
            TxtCant.Text = ""                                   ' Cantidad
            LblUnidad.Caption = ""                              ' UM
            DTPHoras.Value = Format(fIni, "HH:mm")              ' Hora de Inicio
            LblHorFin.Caption = 0                               ' Hora de Fin
            lblntrab.Caption = 0                                ' Numero de trabajadores
            LblDetTrab.Caption = "0 de 0"                       ' Detalle de Trabajadores seleccionados
            TxtIdLineaDet.Text = ""                             ' id de Linea
            LblLinea.Caption = ""                               ' Detalle de Linea
            TxtIdEncarg.Text = ""                               ' idencargado
            LblEncargado.Caption = ""                           ' Encargado
        End If
            
        ' Se Agrega las Tareas
        pCargarDatos fg(0), False, True, , , , DISEÑO_
        calcularDatosAdicionales False
        ' Se Agrega al personal
        pCargarDatos fg(1), True, False, , , , DISEÑO_
        '****************************************************
        ' Se Agrega los Reprocesos
        pCargarDatos fg(4), False, , , , , DISEÑO_, True
        '****************************************************
            
        lblidprocorr.Caption = 0
        lblnumregprod.Caption = ""
        
        '***************************
        txtEfic.Text = 100
        '***************************
                        
        If Not frm(2).Visible Then CentrarFrm frm(2)
        TabOne2.CurrTab = 0
        frm(2).Visible = True
        
        If DISEÑO_ Then
            Cmd(2).SetFocus
        Else
            Cmd(0).SetFocus
        End If
        
        ESTADO_ = ESTADOPENDIENTE_
        llenarEstado True, ESTADO_
    End If
    
    If MODIFICAR_ Then ' Muestra el formulario para modificar productos
        Agregando = True
        
        If DISEÑO_ Then
            IDCRDET_ = fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_)
        Else
            If DETECTOR_.ViewEvent Is Nothing Then Exit Sub
            IDCRDET_ = EVENTO_.ReminderSoundFile
        End If
        
        ' Buscamos el Producto involucrado
        RstProductos.Filter = "id = " & IDCRDET_ & ""
        If RstProductos.RecordCount = 0 Then Exit Sub
        ' Limpiamos el recordset Auxiliar
        limpiarRST RstProductosAux, True
        ' Cargamos el Recordset Auxiliar
        If RstProductosAux.State = 0 Then DEFINIR_RST_TMP RstProductosAux, RstProductos
        CARGAR_RST_TMP RstProductosAux, RstProductos
        
        ' Buscamos la Tarea involucrada
        RstTareas.Filter = "idcrdet = " & IDCRDET_ & ""
        ' Limpiamos el recordset Auxiliar
        limpiarRST RstTareasAux, True
        ' Cargamos el Recordset Auxiliar
        CARGAR_RST_TMP RstTareasAux, RstTareas
        
        If Not DISEÑO_ Then
            LblIdCrDet.Caption = IDCRDET_                                           ' Correlativo
            TxtMatProd.Text = NulosN(RstProductosAux("iditem"))                     ' Id item
            LblMatProd.Caption = NulosC(RstProductosAux("descripcion"))             ' Descripcion Prod
            TxtCodRec.Text = NulosC(RstProductosAux("codrec"))                      ' Codigo de receta
            LblIdRec.Caption = NulosN(RstProductosAux("idrec"))                     ' Id receta
            TxtCant.Text = Format(NulosN(RstProductosAux("cantidad")), "0.00")      ' Cantidad
            LblUnidad.Caption = NulosC(RstProductosAux("abrev"))                    ' UM
            DTPHoras.Value = Format(RstProductosAux("horpro"), "HH:mm")             ' Hora de Inicio
            LblHorFin.Caption = Format(RstProductosAux("horfin"), "HH:mm")          ' Hora de Fin
            lblntrab.Caption = 0                                                    ' Numero de trabajadores
            LblDetTrab.Caption = "0 de 0"                                           ' Detalle de Trabajadores seleccionados
            TxtIdEncarg.Text = NulosN(RstProductosAux("idresp"))
            LblEncargado.Caption = NulosC(RstProductosAux("nomresp"))
            TxtIdLineaDet.Text = NulosN(RstProductosAux("idlinea"))
            LblLinea.Caption = NulosC(RstProductosAux("nomlinea"))
        End If
        
        LblDia.Caption = Format(RstProductosAux("fchpro"), "dd/mm/yyyy")        ' Dia de Programacion
        lblNumprod.Caption = NulosC(RstProductosAux("numprod"))
        
        lblidprocorr.Caption = NulosN(RstProductosAux("idprocorr"))
        lblnumregprod.Caption = NulosC(RstProductosAux("numregprod"))
        
        '*************************************************************
        txtEfic.Text = NulosN(RstProductosAux("efic"))
        '*************************************************************
        
        ESTADO_ = NulosN(RstProductosAux("estado"))
        llenarEstado True, ESTADO_
        
        ' Se Agrega las Tareas
        pCargarDatos fg(0), False, True, , , , DISEÑO_
        calcularDatosAdicionales DISEÑO_
        ' Se Agrega al personal
        pCargarDatos fg(1), True, False, , , , DISEÑO_
        
        '****************************************************
        ' Se Agrega los Reprocesos
        pCargarDatos fg(4), False, , , , , DISEÑO_, True
        '****************************************************
        
        ' Se bloquean los controles
        bloquearControles IDCRDET_, DISEÑO_
        
        If Not frm(2).Visible Then CentrarFrm frm(2)
        TabOne2.CurrTab = 0
        frm(2).Visible = True
        If fg(0).Rows > fg(0).FixedRows Then fg(0).SetFocus: fg_RowColChange 0
        
        Agregando = False
    End If
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

Private Sub Form_Unload(Cancel As Integer)
    If TabOne1.CurrTab = 0 Then Exit Sub
    If Not verificarProcesados Then Cancel = True
End Sub

Private Sub Menu2_1_Click()
    ' Agregar
    bloquearControles 0, False
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

Private Sub cargarSemanasReg()
    Dim xRs As New ADODB.Recordset
    
    cSQL = "SELECT pro_cronograma.id, pro_cronograma.semana " _
                + vbCr + "FROM pro_cronograma " _
                + vbCr + "Where (((pro_cronograma.semana) <> " & NulosN(ComboSemanas.Text) & ")) " _
                + vbCr + "ORDER BY pro_cronograma.semana;"
    
    RST_Busq xRs, cSQL, xCon
    
    cbsemcamb.Clear
    If xRs.State = 0 Then Exit Sub
    xRs.Filter = adFilterNone
    If xRs.RecordCount = 0 Then Exit Sub
    
    xRs.MoveFirst
    While Not xRs.EOF
        'se cargan las semanas
        cbsemcamb.AddItem NulosN(xRs("semana"))
        xRs.MoveNext
    Wend
End Sub

Private Sub cargarDiasSemanaReg(NUMSEM_ As Double)
    Dim FCHINI_ As Date
    Dim FCHFIN_ As Date
    Dim A As Date
    Dim xRs As New ADODB.Recordset
    
    cSQL = "SELECT pro_cronograma.id, pro_cronograma.fchini, pro_cronograma.fchfin " _
        + vbCr + "From pro_cronograma " _
        + vbCr + "Where (((pro_cronograma.semana) = " & NUMSEM_ & ")) " _
        + vbCr + "ORDER BY pro_cronograma.semana;"
    
    RST_Busq xRs, cSQL, xCon
    cbfchcamb.Clear
    If xRs.State = 0 Then Exit Sub
    xRs.Filter = adFilterNone
    If xRs.RecordCount = 0 Then Exit Sub
    
    FCHINI_ = CDate(xRs("fchini"))
    FCHFIN_ = CDate(xRs("fchfin"))
    
    For A = FCHINI_ To FCHFIN_
        cbfchcamb.AddItem Format(A, "dd/mm/yyyy")
    Next A
End Sub

Private Sub menu2_4_Click()
    Dim IDCRDET_ As Double
    Dim DISEÑO_ As Boolean
    Dim DIFDIAS_ As Integer
    
    cbsemcamb.Clear
    cbfchcamb.Clear
    
    If CalCtrlCronog.Visible Then DISEÑO_ = False Else DISEÑO_ = True
        
    If DISEÑO_ Then
        IDCRDET_ = fg(3).TextMatrix(fg(3).Row, COLUMNAIDCRDET_)
        
    Else
        If DETECTOR_.ViewEvent Is Nothing Then
            IDCRDET_ = 0
            GoTo FILTRAR_
        End If
        IDCRDET_ = EVENTO_.ReminderSoundFile
    End If
    
FILTRAR_:
    ' Buscamos el Producto involucrado
    RstProductos.Filter = "id = " & IDCRDET_ & ""
    If RstProductos.RecordCount = 0 Then Exit Sub
    
    DIFDIAS_ = CDate(RstProductos("fchpro")) - CDate(RstProductos("fchfin"))
    
    LblProd.Caption = NulosC(RstProductos("descripcion"))
    LblDetProd(0).Caption = IDCRDET_
    LblDetProd(2).Caption = DIFDIAS_
    cargarSemanasReg
    CentrarFrm frm(1)
    frm(1).Visible = True
End Sub

Private Sub Menu3_1_Click()
    ' Productos de receta
    mostrarFormulario False, False, True
End Sub

Private Sub Menu4_1_Click()
'    Dim FILAINICIO_ As Integer
'    Dim FILAFIN_ As Integer
'    Dim COLUMNAINICIO_ As Integer
'    Dim COLUMNAFIN_ As Integer
'    Dim A As Integer
'    Dim contador As Integer
'
'    hallarRangoSeleccion fg(3), FILAINICIO_, FILAFIN_, COLUMNAINICIO_, COLUMNAFIN_
'
'    If FILAINICIO_ < fg(3).FixedRows Then Exit Sub
'    frm(5).Visible = True
'    centrarFrm frm(5)
'    fg(3).Refresh
'    pgBar.Min = 0
'    pgBar.Max = Abs(FILAFIN_ - FILAINICIO_) + 1
'    contador = 0
'
'    For A = FILAINICIO_ To FILAFIN_
'        fg(3).Select A, COLUMNAINICIO_
'
'        contador = contador + 1
'        frm(5).Refresh
'        pgBar.Value = contador
'        lblProcesado.Caption = fg(3).TextMatrix(A, COLUMNAPRODUCTO_)
'
'        If NulosC(fg(3).TextMatrix(fg(3).Row, COLUMNAFCHPROD_)) = "" Then ' Fecha de Inicio
'            MsgBox "Ingrese Fecha de Programación", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'            fg(3).Select fg(3).Row, COLUMNAFCHPROD_
'            GoTo SIGUIENTE
'        End If
'
'        If NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNAIDITEM_)) = 0 Then ' Producto
'            MsgBox "Ingrese Producto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'            fg(3).Select fg(3).Row, COLUMNAPRODUCTO_
'            GoTo SIGUIENTE
'        End If
'
'        If NulosN(fg(3).TextMatrix(fg(3).Row, COLUMNACANTIDAD_)) = 0 Then ' Cantidad
'            MsgBox "Ingrese Cantidad", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'            fg(3).Select fg(3).Row, COLUMNACANTIDAD_
'            GoTo SIGUIENTE
'        End If
'
'        If fg(3).TextMatrix(fg(3).Row, COLUMNAIDRESP_) = "" Then ' Encargado
'            MsgBox "Ingrese Encargado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'            fg(3).Select fg(3).Row, COLUMNAENCARGADO_
'            GoTo SIGUIENTE
'        End If
'
'        If fg(3).TextMatrix(fg(3).Row, COLUMNAHORINI_) = "" Then ' Hora de Inicio
'            MsgBox "Ingrese Hora de Inicio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'            fg(3).Select fg(3).Row, COLUMNAHORINI_
'            GoTo SIGUIENTE
'        End If
'
'        If procesarLineaProduccion(fg(0), True) Then
'            aplicarCambios True
'            fg(3).Select A, 1, A, fg(3).Cols - 1
'            fg(3).FillStyle = flexFillRepeat
'            fg(3).CellBackColor = &H80000005
'        End If
'
'SIGUIENTE:
'    Next A
'
'    frm(5).Visible = False
End Sub

Private Sub hallarRangoSeleccion(ByRef FGRID_ As VSFlexGrid, ByRef FILAINI_ As Integer, _
                                ByRef FILAFIN_ As Integer, ByRef COLINI_ As Integer, _
                                ByRef COLFIN_ As Integer)
    With FGRID_
        ' Columnas
        If .Col < .ColSel Then COLINI_ = .Col: COLFIN_ = .ColSel
        If .Col > .ColSel Then COLFIN_ = .Col: COLINI_ = .ColSel
        If .Col = .ColSel Then COLINI_ = .Col: COLFIN_ = .Col
        ' Filas
        If .Row < .RowSel Then FILAINI_ = .Row: FILAFIN_ = .RowSel
        If .Row > .RowSel Then FILAFIN_ = .Row: FILAINI_ = .RowSel
        If .Row = .RowSel Then FILAINI_ = .Row: FILAFIN_ = .Row
    End With
End Sub

Private Sub Menu4_2_Click()
    Dim FILAINICIO_ As Integer
    Dim FILAFIN_ As Integer
    Dim COLUMNAINICIO_ As Integer
    Dim COLUMNAFIN_ As Integer
    Dim A As Integer
    Dim B As Integer
    Dim contador As Integer
    Dim dato() As Variant
    
    hallarRangoSeleccion fg(3), FILAINICIO_, FILAFIN_, COLUMNAINICIO_, COLUMNAFIN_
    ReDim dato(COLUMNAINICIO_ To COLUMNAFIN_) As Variant
    
    With fg(3)
        For A = FILAINICIO_ + 1 To FILAFIN_
            For B = COLUMNAINICIO_ To COLUMNAFIN_
                If B = COLUMNAPRODUCTO_ Then ' Producto
                    .TextMatrix(A, COLUMNAIDITEM_) = .TextMatrix(FILAINICIO_, COLUMNAIDITEM_)
                    .TextMatrix(A, COLUMNARECETA_) = .TextMatrix(FILAINICIO_, COLUMNARECETA_)
                    .TextMatrix(A, COLUMNAIDRECETA_) = .TextMatrix(FILAINICIO_, COLUMNAIDRECETA_)
                End If
                
                If B = COLUMNARECETA_ Then ' Receta
                    ' Si se trata del mismo producto
                    If .TextMatrix(A, COLUMNAIDITEM_) = .TextMatrix(FILAINICIO_, COLUMNAIDITEM_) Then
                        .TextMatrix(A, COLUMNAIDRECETA_) = .TextMatrix(FILAINICIO_, COLUMNAIDRECETA_)
                        .TextMatrix(A, B) = .TextMatrix(FILAINICIO_, B)
                    Else
                        .TextMatrix(A, COLUMNAIDRECETA_) = ""
                        .TextMatrix(A, B) = ""
                    End If
                    GoTo SIGUIENTE
                End If
                
                If B = COLUMNAENCARGADO_ Then ' Encargado
                    .TextMatrix(A, COLUMNAIDRESP_) = .TextMatrix(FILAINICIO_, COLUMNAIDRESP_)
                End If
                
                If B = COLUMNALINEA_ Then ' linea
                    ' Si se trata del mismo producto
                    If .TextMatrix(A, COLUMNAIDITEM_) = .TextMatrix(FILAINICIO_, COLUMNAIDITEM_) _
                            And .TextMatrix(A, COLUMNAIDRECETA_) = .TextMatrix(FILAINICIO_, COLUMNAIDRECETA_) Then
                        .TextMatrix(A, COLUMNAIDLINEA_) = .TextMatrix(FILAINICIO_, COLUMNAIDLINEA_)
                        .TextMatrix(A, B) = .TextMatrix(FILAINICIO_, B)
                    Else
                        .TextMatrix(A, COLUMNAIDLINEA_) = ""
                        .TextMatrix(A, B) = ""
                    End If
                    GoTo SIGUIENTE
                End If
                
                .TextMatrix(A, B) = .TextMatrix(FILAINICIO_, B)
SIGUIENTE:
                If B = COLUMNAPROCESADO_ Or B = COLUMNAHORFIN_ _
                                    Or B = COLUMNAFCHFIN_ Or B = COLUMNANUMOPE_ Then ' Procesado
                    .TextMatrix(A, B) = ""
                End If
            Next B
        Next A
    End With
End Sub

Private Sub menu4_3_Click()
    menu2_4_Click
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
        frm(2).Visible = False
    End If
    
    If Index = 1 Then
        frm(1).Visible = False
    End If
    
    If Index = 2 Then
        frm(0).Visible = False
    End If
    
    If Index = 3 Then
        frm(4).Visible = False
    End If
End Sub

'Private Sub SliderCal_Change()
'    CalCtrlCronog.DayView.MinColumnWidth = SliderCal.Value
'    CalCtrlCronog.RedrawControl
'    CalCtrlCronog.Populate
'End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    ' Se limpia el calendario
    'CalCtrlCronog.DataProvider.RemoveAllEvents
    If OldTab = 0 Then
        If QueHace = 1 Then Exit Sub
        MuestraSegundoTab
    Else
        'frm(2).Visible = False
        frm(0).Visible = False
        frm(4).Visible = False
        'CalCtrlCronog.Visible = True
        
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
    
    If Button.Index = 6 Then
        If verificarProcesados Then Cancelar
    End If
    
    If Button.Index = 9 Then
        If TabOne1.CurrTab = 0 Then RstLis.Filter = "": TDB_FiltroLimpiar Dg1
        If TabOne1.CurrTab = 1 Then CmdOpciones_Click 0
    End If
    
    If Button.Index = 12 Then
        If TabOne1.CurrTab = 0 Then Exit Sub
        imprimir 1
    End If
    
    If Button.Index = 14 Then
        Set RstLis = Nothing
        Unload Me
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then ' Imprimir linea
        If TabOne1.CurrTab = 0 Then Exit Sub
        imprimir 0
    End If
    If ButtonMenu.Index = 2 Then ' Imprimir Reporte
        If TabOne1.CurrTab = 0 Then Exit Sub
        imprimir 1
    End If
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

'Private Sub TxtCantMP_KeyPress(KeyAscii As Integer)
'    KeyAscii = 0
'End Sub

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
        cmd_Click 18
    End If
End Sub

Private Sub TxtIdEncarg_Validate(Cancel As Boolean)
    If NulosN(TxtIdEncarg.Text) = 0 Then
        TxtIdEncarg.Text = ""
        LblEncargado.Caption = ""
        Exit Sub
    Else
        LblEncargado.Caption = Busca_Codigo(NulosN(TxtIdEncarg.Text), "id", "nombre", "pla_empleados", "N", xCon)
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
        cmd_Click 0
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
            If frm(2).Visible Then TxtCant.SetFocus
        End If
    End If
End Sub

Private Sub TxtPctje_Change()
    TxtPctje.Text = Format(NulosN(TxtPctje.Text), "0.00")
End Sub

Private Sub TxtPctje_GotFocus()
    Me.TxtPctje.SelStart = 0
    Me.TxtPctje.SelLength = Len(Me.TxtPctje.Text)
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
