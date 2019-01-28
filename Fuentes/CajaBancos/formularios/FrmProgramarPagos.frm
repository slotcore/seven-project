VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmProgramarPagos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caja y Bancos - Programar Pagos"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7245
      Left            =   15
      TabIndex        =   11
      Top             =   375
      Width           =   11880
      _cx             =   20955
      _cy             =   12779
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
         Height          =   6825
         Left            =   12525
         TabIndex        =   15
         Top             =   375
         Width           =   11790
         Begin VB.CommandButton CmdModificar 
            Caption         =   "&Modificar"
            Height          =   375
            Index           =   0
            Left            =   8760
            TabIndex        =   63
            Top             =   1995
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.CommandButton CmdModificar 
            Caption         =   "&Cancelar"
            Height          =   375
            Index           =   1
            Left            =   10215
            TabIndex        =   65
            Top             =   1995
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.TextBox txt 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   1
            Left            =   6945
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "txt(1)"
            Top             =   1815
            Width           =   1545
         End
         Begin VB.TextBox txtApro 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   300
            Index           =   3
            Left            =   10320
            Locked          =   -1  'True
            TabIndex        =   61
            Text            =   "txtApro(3)"
            ToolTipText     =   "Total Aprobado: Nuevo Saldo"
            Top             =   5370
            Width           =   1065
         End
         Begin VB.TextBox txtApro 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   300
            Index           =   0
            Left            =   7035
            Locked          =   -1  'True
            TabIndex        =   60
            Text            =   "txtApro(0)"
            ToolTipText     =   "Total Aprobado: Importe"
            Top             =   5370
            Width           =   1065
         End
         Begin VB.TextBox txtApro 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   300
            Index           =   1
            Left            =   8130
            Locked          =   -1  'True
            TabIndex        =   59
            Text            =   "txtApro(1)"
            ToolTipText     =   "Total Aprobado: Saldo Anterior"
            Top             =   5370
            Width           =   1065
         End
         Begin VB.TextBox txtApro 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   300
            Index           =   2
            Left            =   9225
            Locked          =   -1  'True
            TabIndex        =   58
            Text            =   "txtApro(2)"
            ToolTipText     =   "Total Aprobado: A Cuenta"
            Top             =   5370
            Width           =   1065
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Index           =   3
            Left            =   10320
            Locked          =   -1  'True
            TabIndex        =   56
            Text            =   "txtTotal(3)"
            ToolTipText     =   "Total: Nuevo Saldo"
            Top             =   5055
            Width           =   1065
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   7035
            Locked          =   -1  'True
            TabIndex        =   55
            Text            =   "txtTotal(0)"
            ToolTipText     =   "Total: Importe"
            Top             =   5055
            Width           =   1065
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   8130
            Locked          =   -1  'True
            TabIndex        =   54
            Text            =   "txtTotal(1)"
            ToolTipText     =   "Total: Saldo Anterior"
            Top             =   5055
            Width           =   1065
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Index           =   2
            Left            =   9225
            Locked          =   -1  'True
            TabIndex        =   53
            Text            =   "txtTotal(2)"
            ToolTipText     =   "Total: A Cuenta"
            Top             =   5055
            Width           =   1065
         End
         Begin VB.TextBox txtRech 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Index           =   3
            Left            =   10320
            Locked          =   -1  'True
            TabIndex        =   50
            Text            =   "txtRech(3)"
            ToolTipText     =   "Total Rechazado: Nuevo Saldo"
            Top             =   5685
            Width           =   1065
         End
         Begin VB.TextBox txtRech 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Index           =   0
            Left            =   7035
            Locked          =   -1  'True
            TabIndex        =   49
            Text            =   "txtRech(0)"
            ToolTipText     =   "Total Rechazado: Importe"
            Top             =   5685
            Width           =   1065
         End
         Begin VB.TextBox txtRech 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Index           =   1
            Left            =   8130
            Locked          =   -1  'True
            TabIndex        =   48
            Text            =   "txtRech(1)"
            ToolTipText     =   "Total Rechazado: Saldo Anterior"
            Top             =   5685
            Width           =   1065
         End
         Begin VB.TextBox txtRech 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Index           =   2
            Left            =   9225
            Locked          =   -1  'True
            TabIndex        =   47
            Text            =   "txtRech(2)"
            ToolTipText     =   "Total Rechazado: A Cuenta"
            Top             =   5685
            Width           =   1065
         End
         Begin VB.Frame Frame3 
            Caption         =   "( Periodo )"
            Height          =   720
            Left            =   9645
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
         Begin VB.Frame fra_estado 
            Height          =   1080
            Left            =   8700
            TabIndex        =   28
            Top             =   885
            Width           =   2955
            Begin VB.CommandButton cmd_estado 
               Caption         =   "&Rechazar"
               Height          =   315
               Index           =   1
               Left            =   1515
               TabIndex        =   34
               Top             =   690
               Width           =   1365
            End
            Begin VB.CommandButton cmd_estado 
               Caption         =   "&Aprobar"
               Height          =   315
               Index           =   0
               Left            =   105
               TabIndex        =   33
               Top             =   690
               Width           =   1365
            End
            Begin VB.Label LblEstado 
               Alignment       =   2  'Center
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   465
               Index           =   1
               Left            =   2385
               TabIndex        =   32
               Top             =   180
               Visible         =   0   'False
               Width           =   510
            End
            Begin VB.Label LblEstado 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Pendiente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   450
               Index           =   0
               Left            =   120
               TabIndex        =   29
               Top             =   195
               Width           =   2775
            End
         End
         Begin VB.TextBox txt 
            Height          =   705
            Index           =   2
            Left            =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Tag             =   "null"
            Text            =   "FrmProgramarPagos.frx":0000
            Top             =   6060
            Width           =   11430
         End
         Begin VB.Frame fr 
            Caption         =   "[ Tipo de Operación ]"
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
            Height          =   630
            Index           =   0
            Left            =   5985
            TabIndex        =   22
            Top             =   885
            Width           =   2505
            Begin VB.OptionButton opt_operacion 
               Caption         =   "Banco"
               Height          =   195
               Index           =   1
               Left            =   1290
               TabIndex        =   10
               Top             =   330
               Width           =   840
            End
            Begin VB.OptionButton opt_operacion 
               Caption         =   "Caja"
               Height          =   195
               Index           =   0
               Left            =   180
               TabIndex        =   3
               Top             =   330
               Value           =   -1  'True
               Width           =   840
            End
         End
         Begin VB.CommandButton cb 
            Height          =   225
            Index           =   0
            Left            =   1845
            Picture         =   "FrmProgramarPagos.frx":0009
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1845
            Width           =   210
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   0
            Left            =   1410
            MaxLength       =   10
            TabIndex        =   2
            Text            =   "txt_cb(0)"
            Top             =   1815
            Width           =   675
         End
         Begin VB.TextBox txt 
            BackColor       =   &H0080FF80&
            Height          =   315
            Index           =   0
            Left            =   7665
            TabIndex        =   18
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   330
            Visible         =   0   'False
            Width           =   1170
         End
         Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
            Height          =   300
            Index           =   0
            Left            =   1410
            TabIndex        =   0
            Tag             =   "b"
            Top             =   1140
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
         End
         Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
            Height          =   300
            Index           =   1
            Left            =   1410
            TabIndex        =   1
            Tag             =   "b"
            Top             =   1470
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
         End
         Begin VSFlex7Ctl.VSFlexGrid fg1 
            Height          =   2625
            Left            =   150
            TabIndex        =   8
            Top             =   2415
            Width           =   11505
            _cx             =   20294
            _cy             =   4630
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
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   14
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmProgramarPagos.frx":013B
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
         Begin VB.Frame Frame4 
            Height          =   870
            Left            =   150
            TabIndex        =   44
            Top             =   4980
            Width           =   3900
            Begin VB.CommandButton Cmd 
               Caption         =   "&Seleccionar Documentos"
               Enabled         =   0   'False
               Height          =   540
               Index           =   1
               Left            =   1410
               Style           =   1  'Graphical
               TabIndex        =   6
               Tag             =   "b"
               Top             =   195
               Width           =   1170
            End
            Begin VB.CommandButton Cmd 
               Caption         =   "&Agregar Documentos"
               Enabled         =   0   'False
               Height          =   540
               Index           =   0
               Left            =   210
               Style           =   1  'Graphical
               TabIndex        =   5
               Tag             =   "b"
               Top             =   195
               Width           =   1170
            End
            Begin VB.CommandButton Cmd 
               Caption         =   "&Eliminar Documento"
               Enabled         =   0   'False
               Height          =   540
               Index           =   2
               Left            =   2625
               Style           =   1  'Graphical
               TabIndex        =   7
               Tag             =   "b"
               Top             =   195
               Width           =   1170
            End
         End
         Begin VB.Frame Frame5 
            Height          =   870
            Left            =   4050
            TabIndex        =   45
            Top             =   4980
            Width           =   1035
            Begin VB.CheckBox ChkAutorizar 
               Caption         =   "Aprobar Todos"
               Enabled         =   0   'False
               Height          =   405
               Left            =   75
               TabIndex        =   46
               Top             =   300
               Width           =   900
            End
         End
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6555
            TabIndex        =   57
            Top             =   5160
            Width           =   435
         End
         Begin VB.Label lbl_aut 
            BackColor       =   &H0000FFFF&
            Caption         =   "lbl_aut(0)"
            Height          =   270
            Index           =   0
            Left            =   4695
            TabIndex        =   27
            Top             =   765
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "N°. Documento"
            Height          =   195
            Index           =   1
            Left            =   5730
            TabIndex        =   62
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label LblTotAprob 
            AutoSize        =   -1  'True
            Caption         =   "Total Aprobados"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5595
            TabIndex        =   52
            Top             =   5475
            Width           =   1395
         End
         Begin VB.Label LblTotRechaz 
            AutoSize        =   -1  'True
            Caption         =   "Total Rechazados"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5490
            TabIndex        =   51
            Top             =   5790
            Width           =   1500
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Documentos a Pagar"
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
            Index           =   1
            Left            =   150
            TabIndex        =   43
            Top             =   2190
            Width           =   1785
         End
         Begin VB.Label LblTipCam2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Cambio"
            Height          =   195
            Left            =   2880
            TabIndex        =   42
            Top             =   1200
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
            Left            =   4065
            TabIndex        =   41
            Top             =   1095
            Width           =   1350
         End
         Begin VB.Label lbl_aut 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Autorizador:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   37
            Top             =   810
            Width           =   840
         End
         Begin VB.Label lbl_prog 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Programador:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   36
            Top             =   495
            Width           =   945
         End
         Begin VB.Label lbl_aut 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_aut(1)"
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
            Left            =   1410
            TabIndex        =   30
            Top             =   765
            Width           =   4365
         End
         Begin VB.Label lbl_prog 
            BackColor       =   &H0000FFFF&
            Caption         =   "lbl_prog(0)"
            Height          =   270
            Index           =   0
            Left            =   4695
            TabIndex        =   26
            Top             =   450
            Visible         =   0   'False
            Width           =   855
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
            Left            =   3675
            TabIndex        =   25
            Top             =   1815
            Visible         =   0   'False
            Width           =   1230
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
            Left            =   2070
            TabIndex        =   24
            Top             =   1815
            Width           =   1830
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Observación"
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   23
            Top             =   5850
            Width           =   900
         End
         Begin VB.Label lblfch 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Pago"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   21
            Top             =   1545
            Width           =   870
         End
         Begin VB.Label lblfch 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Emisión"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   20
            Top             =   1185
            Width           =   1035
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   19
            Top             =   1905
            Width           =   585
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "codigo"
            Height          =   195
            Index           =   0
            Left            =   7065
            TabIndex        =   17
            Top             =   450
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Programación de Pagos"
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
            Left            =   0
            TabIndex        =   16
            Top             =   30
            Width           =   11550
         End
         Begin VB.Label lbl_prog 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_prog(1)"
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
            Left            =   1410
            TabIndex        =   31
            Top             =   450
            Width           =   4365
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6825
         Left            =   45
         TabIndex        =   13
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg3 
            Height          =   6465
            Left            =   15
            TabIndex        =   64
            Top             =   345
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   11404
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
            Columns(1).Caption=   "Operación"
            Columns(1).DataField=   "Operacion"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Emi."
            Columns(2).DataField=   "fchemi"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Fch.Pago"
            Columns(3).DataField=   "fchpag"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "N° Doc"
            Columns(4).DataField=   "numdoc"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "M"
            Columns(5).DataField=   "simbolo"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Imp. Total"
            Columns(6).DataField=   "imptot"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Imp.Acep."
            Columns(7).DataField=   "impapro"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Imp.Rech"
            Columns(8).DataField=   "imprech"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Emitido por"
            Columns(9).DataField=   "prog"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Autorizado por"
            Columns(10).DataField=   "aut"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Estado"
            Columns(11).DataField=   "estdesc"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   12
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=12"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1058"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=979"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=514"
            Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(7)=   "Column(1).Width=1667"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1588"
            Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(13)=   "Column(2).Width=1561"
            Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1482"
            Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(19)=   "Column(3).Width=1614"
            Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1535"
            Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(25)=   "Column(4).Width=2090"
            Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2011"
            Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(31)=   "Column(5).Width=794"
            Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=714"
            Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(37)=   "Column(6).Width=1879"
            Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=1799"
            Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=514"
            Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(43)=   "Column(7).Width=1720"
            Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=1640"
            Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=514"
            Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(49)=   "Column(8).Width=1746"
            Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1667"
            Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=514"
            Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(55)=   "Column(9).Width=1826"
            Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=1746"
            Splits(0)._ColumnProps(58)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(59)=   "Column(9)._ColStyle=516"
            Splits(0)._ColumnProps(60)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(61)=   "Column(10).Width=2170"
            Splits(0)._ColumnProps(62)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(63)=   "Column(10)._WidthInPix=2090"
            Splits(0)._ColumnProps(64)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(65)=   "Column(10)._ColStyle=516"
            Splits(0)._ColumnProps(66)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(67)=   "Column(11).Width=1693"
            Splits(0)._ColumnProps(68)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(69)=   "Column(11)._WidthInPix=1614"
            Splits(0)._ColumnProps(70)=   "Column(11)._EditAlways=0"
            Splits(0)._ColumnProps(71)=   "Column(11)._ColStyle=516"
            Splits(0)._ColumnProps(72)=   "Column(11).Order=12"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HDBFDFD&,.bold=0,.fontsize=825"
            _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
            _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.bgcolor=&H80&"
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
            _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38,.bgcolor=&H80&"
            _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=62,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=90,.parent=13,.alignment=1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=87,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=88,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=89,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=43,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=44,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=45,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=32,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=29,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=30,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=31,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
            _StyleDefs(84)  =   "Named:id=33:Normal"
            _StyleDefs(85)  =   ":id=33,.parent=0"
            _StyleDefs(86)  =   "Named:id=34:Heading"
            _StyleDefs(87)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(88)  =   ":id=34,.wraptext=-1"
            _StyleDefs(89)  =   "Named:id=35:Footing"
            _StyleDefs(90)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(91)  =   "Named:id=36:Selected"
            _StyleDefs(92)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(93)  =   "Named:id=37:Caption"
            _StyleDefs(94)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(95)  =   "Named:id=38:HighlightRow"
            _StyleDefs(96)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(97)  =   "Named:id=39:EvenRow"
            _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(99)  =   "Named:id=40:OddRow"
            _StyleDefs(100) =   ":id=40,.parent=33"
            _StyleDefs(101) =   "Named:id=41:RecordSelector"
            _StyleDefs(102) =   ":id=41,.parent=34"
            _StyleDefs(103) =   "Named:id=42:FilterBar"
            _StyleDefs(104) =   ":id=42,.parent=33"
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
            TabIndex        =   38
            Top             =   75
            Width           =   1980
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Programar Pagos"
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
            Left            =   15
            TabIndex        =   14
            Top             =   30
            Width           =   11550
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
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
         Left            =   5535
         Top             =   60
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
               Picture         =   "FrmProgramarPagos.frx":02D7
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramarPagos.frx":081B
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramarPagos.frx":0BAD
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramarPagos.frx":0D31
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramarPagos.frx":1185
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramarPagos.frx":129D
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramarPagos.frx":17E1
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramarPagos.frx":1D25
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramarPagos.frx":1E39
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramarPagos.frx":1F4D
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramarPagos.frx":23A1
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmProgramarPagos.frx":250D
               Key             =   "IMG11"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Agregar Compra"
      End
      Begin VB.Menu menu1_4 
         Caption         =   "Seleccionar Compra"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar Compra"
      End
   End
End
Attribute VB_Name = "FrmProgramarPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim RstFrm As New ADODB.Recordset
Dim Agregando As Boolean

Dim fOcultarToolbar As Boolean  '--FALSE::SE OCULTA TRUE::MOSTRAR

Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta


'--de los estados
'LblEstado(1).Caption = "2"
'LblEstado(1).Caption = "4"
'

Private Sub cb_Click(Index As Integer)
    Dim xCampos() As String
    Dim nTitulo As String
    Dim nOrden As String
    Dim nCampoBusca As String
    Dim nSQL As String
    
    If QueHace = 3 Then Exit Sub
    
    On Error GoTo error
    Select Case Index
    
        Case 0 '--MONEDA
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Moneda":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "3500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Símbolo":   xCampos(1, 1) = "simbolo":    xCampos(1, 2) = "1000":   xCampos(1, 3) = "C"
            nTitulo = "Buscando Moneda"
            nSQL = "SELECT mae_moneda.id, mae_moneda.descripcion as nombre,mae_moneda.id as cod,mae_moneda.simbolo  " _
                + vbCr + " From mae_moneda "
                
            nCampoBusca = "nombre"

    End Select
    nOrden = "nombre"
    
    Dim xRs As New ADODB.Recordset

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, nOrden, nCampoBusca, Principio
    
    If xRs.State = 0 Then GoTo Salir
    If xRs.RecordCount = 0 Then GoTo Salir
    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cb_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
    
    txt(1).SetFocus
    
Salir:
    Set xRs = Nothing
Exit Sub

error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click (" & Index & ")"
End Sub



Private Sub ChkAutorizar_Click()
    Dim A&
    If ChkAutorizar.Value = 1 Then
        For A = 1 To Fg1.Rows - 1
            Fg1.TextMatrix(A, 2) = -1
        Next
    Else
        For A = 1 To Fg1.Rows - 1
            Fg1.TextMatrix(A, 2) = 0
        Next
    End If
End Sub
Private Sub Dg3_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg3.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub




Private Sub Fg1_DblClick()
    '--ver el detalle del documento por compras
    If Fg1.Rows = 1 Then Exit Sub
    If Fg1.Row < 1 Then Exit Sub
    
'*******************
''On Error GoTo error
''    Dim xForm As New sgi2_compras.Compras
''    D = xForm.RegCompras(xCon, xMes, 1, RstFrm("id"))
''    Set xForm = Nothing
''    Exit Sub
''error:
''    Agregando = False
''    Set xForm = Nothing
'**************************

End Sub

Private Sub lbl_cb_cod_Change(Index As Integer)
    If QueHace = 3 Then Exit Sub
    If lbl_cb_cod(Index) <> "" Then SendKeys vbTab
End Sub

Private Sub cmd_estado_Click(Index As Integer)

    If QueHace = 1 Then
        MsgBox "Primero guarde el registro" + vbCr + "Luego proceda a " + cmd_estado(Index).Caption, vbInformation, xTitulo
        Exit Sub
    Else
        If Fg1.Rows = 1 Then
            MsgBox "Ingrese por lo menos un Registro de Compra para Programar Pagos", vbExclamation, xTitulo
            CmdModificar(0).SetFocus
        End If
        
        If NulosN(txtApro(2).Text) = 0 Then
            MsgBox "Active la Opción Aprobar [1ra Columna] por lo menos a un Registro de Compra" + vbCr + "Si desea aprobar todos los registros de compras, puede seleccionar [Activar Todos] ", vbExclamation, xTitulo
            CmdModificar(0).SetFocus
            Exit Sub
        End If
    End If
    If RstFrm.EOF = True Or RstFrm.BOF = True Then
        MsgBox "No hay registros", vbExclamation, xTitulo
        Exit Sub
    End If
    
    If IsNull(RstFrm.Fields("idaut")) = False Then
        If NulosN(RstFrm.Fields("idaut")) <> 0 And NulosN(RstFrm.Fields("idaut")) <> NulosN(lbl_aut(0).Caption) Then
            If MsgBox("Este registro ha sido autorizado por otra persona " + vbCr + "Autorizador: " + RstFrm.Fields("aut") & "" + vbCr + "Desea continuar", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
        End If
    End If
    
    If MsgBox("Seguro desea " + cmd_estado(Index).Caption, vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
        
    Dim RstCab As New ADODB.Recordset
    Dim xId As Double
    On Error GoTo LaCague
    xCon.BeginTrans
    
    xId = NulosN(RstFrm("id"))
    
    RST_Busq RstCab, "select * from con_ordenpago where id = " & xId, xCon
    
    
    '************
    If Index = 0 Then
        PONER_COLOR_ESTADO xCon, LblEstado(0), 2 '--APROBADO
        LblEstado(1).Caption = "2"
        '--deshabilitar boton de mofificar,cancelar
        habilitar CmdModificar, False
    ElseIf Index = 1 Then
        PONER_COLOR_ESTADO xCon, LblEstado(0), 4 '--RECHAZADO
        LblEstado(1).Caption = "4"
        '--habilitar boton de mofificar
        CmdModificar(0).Enabled = True
    End If
    '************
    
    If Index = 0 Then '--APROBAR
        RstCab("idaut") = NulosN(lbl_aut(0).Caption)
    Else
        RstCab("idaut") = 0
    End If
    RstCab("idest") = NulosN(LblEstado(1).Caption)
    
    RstCab.Update
    xCon.CommitTrans
    
    MsgBox "El registro fue " + LblEstado(0).Caption + " con éxito", vbInformation, xTitulo
    
    Me.MousePointer = vbHourglass
    RstFrm.Requery
    Me.MousePointer = vbDefault
    Set RstCab = Nothing
    Exit Sub
LaCague:
    Set RstCab = Nothing
    Me.MousePointer = vbDefault
    xCon.RollbackTrans
    MsgBox "No se pudo modificar el registro por el siguiente motivo :" + Trim(Err.Description), vbCritical
    Exit Sub
error:
    Set RstCab = Nothing
    SHOW_ERROR
End Sub

Private Sub Dg3_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    Dim Rpta As Integer
    
    SeEjecuto = False
    pCargarGrid
    SeEjecuto = True
    If RstFrm.RecordCount = 0 Then
        If fOcultarToolbar = False Then Exit Sub
        If MsgBox("No se ha registrado ninguna cuenta por rendir, ¿Desea agergar uno ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo) = vbYes Then
            Nuevo
        End If
    End If
End Sub

Private Sub pCargarGrid()
    Dim nSQL  As String
    
    LblPeriodo(0).Caption = Busca_Codigo(xMes, "id", "descripcion", "con_meses", "N", xCon)
    LblPeriodo(1).Caption = LblPeriodo(0).Caption
    'nSQL = "SELECT con_ordenpago.*, pla_empleados.nom & ' ' & pla_empleados.ape AS prog, pla_empleados_1.ape & ' ' & pla_empleados_1.nom AS aut, mae_estados.descripcion AS estdesc, IIf(con_ordenpago.tipope=1,'Caja','Banco') AS operacion , mae_moneda.simbolo " _
        + vbCr + " FROM pla_empleados RIGHT JOIN (mae_moneda RIGHT JOIN (con_emptes RIGHT JOIN (mae_estados RIGHT JOIN ((con_ordenpago LEFT JOIN con_emptes AS con_emptes_1 ON con_ordenpago.idaut = con_emptes_1.id) LEFT JOIN pla_empleados AS pla_empleados_1 ON con_emptes_1.idemp = pla_empleados_1.id) ON mae_estados.id = con_ordenpago.idest) ON con_emptes.id = con_ordenpago.idprog) ON mae_moneda.id = con_ordenpago.idmon) ON pla_empleados.id = con_emptes.idemp " _
        + vbCr + " WHERE con_ordenpago.ano = " & AnoTra & " And con_ordenpago.idmes = " & xMes & " ; "
    nSQL = "SELECT con_ordenpago.*, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ', ' & [pla_empleados].[nom] AS prog, " _
        & " [pla_empleados_1].[apepat] & ' ' & [pla_empleados_1].[apemat] & ', ' & [pla_empleados_1].[nom] AS aut, mae_estados.descripcion AS estdesc, " _
        & " IIf(con_ordenpago.tipope=1,'Caja','Banco') AS operacion, mae_moneda.simbolo FROM mae_moneda RIGHT JOIN ((pla_empleados RIGHT JOIN " _
        & " con_emptes ON pla_empleados.id = con_emptes.idemp) RIGHT JOIN (mae_estados RIGHT JOIN ((con_ordenpago LEFT JOIN con_emptes AS con_emptes_1 " _
        & " ON con_ordenpago.idaut = con_emptes_1.id) LEFT JOIN pla_empleados AS pla_empleados_1 ON con_emptes_1.idemp = pla_empleados_1.id) " _
        & " ON mae_estados.id = con_ordenpago.idest) ON con_emptes.id = con_ordenpago.idprog) ON mae_moneda.id = con_ordenpago.idmon " _
        & " WHERE (((con_ordenpago.ano)=2009) AND ((con_ordenpago.idmes)=10))"

    TabOne1.CurrTab = 0
    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQL, xCon
    
    Set Dg3.DataSource = RstFrm
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    CentrarFrm Me
    SeEjecuto = False
    Agregando = False
    QueHace = 3
    Dg3.BatchUpdates = False
    
    Dg3.Columns("fchemi").NumberFormat = FORMAT_DATE
    Dg3.Columns("fchpag").NumberFormat = FORMAT_DATE
    Dg3.Columns("imptot").NumberFormat = FORMAT_MONTO
    Dg3.Columns("impapro").NumberFormat = FORMAT_MONTO
    Dg3.Columns("imprech").NumberFormat = FORMAT_MONTO
   
    Fg1.ColFormat(7) = FORMAT_DATE
    Fg1.ColFormat(8) = FORMAT_DATE
    Fg1.ColFormat(10) = FORMAT_MONTO
    Fg1.ColFormat(11) = FORMAT_MONTO
    Fg1.ColFormat(12) = FORMAT_MONTO
    Fg1.ColFormat(13) = FORMAT_MONTO
   
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    '--SI NO ES PROGRAMADOR
    If fVerificarProgAut(True, lbl_prog(0), lbl_prog(1)) = False Then    '--PROGAMADOR
        Toolbar1.Buttons(1).Enabled = False
        Toolbar1.Buttons(2).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        fOcultarToolbar = False
    Else
        fOcultarToolbar = True
    End If
    '--
    Habilitar_Obj False
    '--AUTORIZADOR
    If fVerificarProgAut(False, lbl_aut(0), lbl_aut(1)) = True Then
        habilitar cmd_estado, True
        Ocultar CmdModificar, True
        CmdModificar(0).Enabled = True
        CmdModificar(1).Enabled = False
'        ChkAutorizar.Enabled = True
    Else
        Ocultar CmdModificar, False
        habilitar cmd_estado, False
'        ChkAutorizar.Enabled = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation, xTitulo
        Cancel = 1
        Exit Sub
    End If
    Set Dg3.DataSource = Nothing
End Sub






Private Sub Menu1_1_Click()
    pRegistroAdd False
End Sub

Private Sub menu1_3_Click()
    pRegistroDel
End Sub

Private Sub menu1_4_Click()
    pRegistroAdd True
End Sub

Private Sub opt_operacion_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then
        txt(1).SetFocus
    End If
End Sub


Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 3 Then MuestraSegundoTab
    ElseIf OldTab = 1 Then
        QueHace = 3
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    If Button.Index = 2 Then Modificar
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            If RstFrm.State = 0 Then Exit Sub
            RstFrm.Requery
            Dg3.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    If Button.Index = 8 Then Filtrar
    If Button.Index = 9 Then
        RstFrm.Filter = ""
        If RstFrm.State = 0 Then Exit Sub
    End If
    If Button.Index = 10 Then Buscar
    If Button.Index = 11 Then CambiarMes
        
    If Button.Index = 15 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

Sub Eliminar()
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.BOF = True Or RstFrm.EOF = True Or RstFrm.RecordCount = 0 Then Exit Sub

    If RstFrm.Fields("idest") <> "1" Then
        MsgBox "No puede Eliminar" + vbCr + "Ya fue " + LblEstado(0).Caption, vbInformation, xTitulo
        Exit Sub
    End If
    
    Dim Rpta As Integer
    
    Rpta = MsgBox("¿Esta seguro de eliminar El registro seleccionado?", vbQuestion + vbYesNo, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETe * FROM con_ordenpagodet WHERE idord = " & RstFrm("id") & ""
        xCon.Execute "DELETe * FROM con_ordenpago WHERE id = " & RstFrm("id") & ""
        MsgBox "El registro fue eliminado con éxito", vbInformation, xTitulo
        RstFrm.Requery
        TabOne1.CurrTab = 0
        Dg3.Refresh
        If RstFrm.RecordCount = 0 Then
            Rpta = MsgBox("No hay ningún registro, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo)
            If Rpta = vbYes Then
                Nuevo
            Else
                Set RstFrm = Nothing
                Unload Me
                Exit Sub
            End If
        End If
    End If
End Sub

Sub Cancelar()
    QueHace = 3
    TabOne1.TabEnabled(0) = True
    ActivaTool
    Habilitar_Obj False
    Label1.Caption = "Detalle del la Cuenta por Rendir"
    Fg1.SelectionMode = flexSelectionByRow
    TabOne1.CurrTab = 0
    TabOne1.CurrTab = 0
    '-----
    fra_estado.Visible = True
    Ocultar CmdModificar, True
    '------
    Dg3.SetFocus
End Sub

Private Sub Modificar()

    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros", vbExclamation, xTitulo
        Exit Sub
    End If
    
    If NulosN(RstFrm.Fields("idest")) = 2 Then
        MsgBox "No puede Modificar" + vbCr + "Ya fue Aprobado", vbInformation, xTitulo
        Exit Sub
    End If
    '--VER SI EL PROGRAMADOR ES EL MISMO AL QUE CREO EL  REGISTRO
    fVerificarProgAut True, lbl_prog(0), lbl_prog(1)
    If RstFrm.Fields("idprog") & "" <> lbl_prog(0).Caption Then
        MsgBox "Ust. no ha Programado este registro" + vbCr + "Sólo puede modificar quien lo programó", vbInformation, xTitulo
        Exit Sub
    End If
    '------
    
    QueHace = 2
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool

    Habilitar_Obj True
    MuestraSegundoTab
    Fg1.SelectionMode = flexSelectionFree
    Label1.Caption = "Modificando Cuentas por Rendir"
    
    fVerificarProgAut True, lbl_prog(0), lbl_prog(1)   '--PROGRAMADOR
    lbl_aut(0).Caption = ""
    lbl_aut(1).Caption = ""
    '-----
    fra_estado.Visible = False
    Ocultar CmdModificar, False
    '------
    TxtFecha(0).SetFocus
End Sub

Private Sub MuestraSegundoTab()
    Dim QueHaceTmp As Integer
    On Error GoTo error
    With RstFrm
        Blanquea
        If .State = 0 Then Exit Sub
        If .EOF = True Or .BOF = True Then Exit Sub
        QueHaceTmp = QueHace
        QueHace = -1
        txt(0).Text = NulosN(.Fields("id")) '--CODIGO
        TxtFecha(0).Valor = .Fields("fchemi")  '--FECHA DE EMISION
        txtfecha_Validate 0, True
        TxtFecha(1).Valor = .Fields("fchpag")  '--FECHA DE PAGO
        '--TIPO DE OPERACION
        If LCase(.Fields("tipope")) = "1" Then
            Me.opt_operacion(0).Value = True '--ES CAJA
        Else
            Me.opt_operacion(1).Value = True '--ES BANCO
        End If
        
        If NulosN(.Fields("idmon")) <> 0 Then
            txt_cb(0).Text = NulosN(.Fields("idmon"))
            txt_cb_Validate 0, False
        End If
        
        txt(1).Text = NulosC(.Fields("numdoc"))
        txt(2).Text = NulosC(.Fields("obs"))
        
        LblEstado(1).Caption = NulosN(.Fields("idest"))
        If NulosN(.Fields("idest")) <> 0 Then
            PONER_COLOR_ESTADO xCon, LblEstado(0), CInt(.Fields("idest"))
            habilitar cmd_estado, True
            If LblEstado(1).Caption = "2" Then '--APROBADO
                '--deshabilitar boton de mofificar,cancelar
                habilitar CmdModificar, False
            Else
                '--habilitar boton de mofificar
                CmdModificar(0).Enabled = True
            End If
            
        Else
            LblEstado(0).Caption = ""
            LblEstado(0).ForeColor = vbBlack
        End If
        '---DEL PROGRAMADOR
        lbl_prog(0).Caption = NulosC(.Fields("idprog"))
        lbl_prog(1).Caption = NulosC(.Fields("prog"))
    End With
    
    If cmd_estado(0).Enabled = True Then fVerificarProgAut False, lbl_aut(0), lbl_aut(1)
    If CmdModificar(0).Caption = "&Grabar" Then
        CmdModificar(0).Caption = "&Modificar"
        CmdModificar(0).Enabled = True
    End If
    CmdModificar(1).Enabled = False
    QueHace = QueHaceTmp
    MuestraDetalle
    
    Exit Sub
error:
    QueHace = QueHaceTmp
    SHOW_ERROR Me.Name, "MuestraSegundoTab"
End Sub

Private Sub MuestraDetalle()
    Dim xRs As New ADODB.Recordset
    Dim A&
    Dim nSQL As String
    On Error GoTo error
    nSQL = fGenerarConsulta(False)
    
    RST_Busq xRs, nSQL, xCon
    Fg1.Rows = 1
    If xRs.State = 0 Then GoTo Salir
    If xRs.RecordCount = 0 Then GoTo Salir
    Agregando = True
    xRs.MoveFirst
    With Fg1
        Do While Not xRs.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = NulosC(xRs.Fields("idcom"))
            .TextMatrix(.Rows - 1, 2) = NulosN(xRs.Fields("aprobado")) '--autorizado
            .TextMatrix(.Rows - 1, 3) = NulosC(xRs.Fields("registro"))
            .TextMatrix(.Rows - 1, 4) = NulosC(xRs.Fields("abrev"))
            .TextMatrix(.Rows - 1, 5) = NulosC(xRs.Fields("simbolo"))
            .TextMatrix(.Rows - 1, 6) = NulosC(xRs.Fields("doc"))
            .TextMatrix(.Rows - 1, 7) = NulosC(xRs.Fields("fchdoc"))
            .TextMatrix(.Rows - 1, 8) = NulosC(xRs.Fields("fchven"))
            .TextMatrix(.Rows - 1, 9) = NulosC(xRs.Fields("nombre"))
            .TextMatrix(.Rows - 1, 10) = NulosN(xRs.Fields("imptot"))
            .TextMatrix(.Rows - 1, 11) = NulosN(xRs.Fields("saldo"))
            .TextMatrix(.Rows - 1, 12) = NulosN(xRs.Fields("acuenta"))
            .TextMatrix(.Rows - 1, 13) = NulosN(xRs.Fields("nuevosaldo"))
            xRs.MoveNext
        Loop
    End With

Salir:
    '--calcular los totales
    pTotalizarDatos
    Set xRs = Nothing
    Agregando = False
    Exit Sub
error:
    Set xRs = Nothing
    Agregando = False
    SHOW_ERROR
End Sub


Private Sub Habilitar_Obj(band As Boolean)
    habilitar_Locked TxtFecha, Not band
    habilitar_Locked txt, Not band
    habilitar_Locked txt_cb, Not band
    habilitar Me.opt_operacion, band
    habilitar Cmd, band
End Sub

Private Sub Blanquea()

    LblTipoCambio.Caption = ""
    LimpiaText Me.TxtFecha
    LimpiaText txt
    LimpiaText lbl_cb
    LimpiaText lbl_cb
    LimpiaText txt_cb
    LimpiaText txtTotal
    LimpiaText txtApro
    LimpiaText txtRech
    Fg1.Rows = 1
End Sub

Private Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub Nuevo()
    QueHace = 1
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    '---
    TxtFecha(0).Valor = Date
    TxtFecha(1).Valor = Date
    '---
    Blanquea
    Habilitar_Obj True
    Label1.Caption = "Programando Pago a Proveedores"
    '-----
    PONER_COLOR_ESTADO xCon, LblEstado(0), 1 'PENDIENTE
    LblEstado(1).Caption = "1"
    '--
    opt_operacion(0).Value = True
    '--CARGAR PROGRAMADOR
    fVerificarProgAut True, lbl_prog(0), lbl_prog(1)   '--PROGRAMADOR
    lbl_aut(0).Caption = ""
    lbl_aut(1).Caption = ""
    '-----
    fra_estado.Visible = False
    Ocultar CmdModificar, False
    '------
    Fg1.Editable = flexEDKbdMouse
    Fg1.SelectionMode = flexSelectionFree
    '-----
    TxtFecha(0).SetFocus
End Sub

Function Grabar() As Boolean
   If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " el registro", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo Salir
    
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xCod As Integer
    Dim xCol, xFil As Integer
    
    On Error GoTo LaCague

    xCon.BeginTrans
    
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT top 1 * FROM con_ordenpago", xCon
        xCod = HallaCodigoTabla("con_ordenpago", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xCod
        txt(0).Text = xCod
    Else
        xCod = RstFrm("id")
        RST_Busq RstCab, "SELECT * FROM con_ordenpago WHERE id =" & xCod & "", xCon
        
        xCon.Execute "Delete From con_ordenpagodet Where idord= " & xCod & ";"
    End If
    '********************************
    RST_Busq RstDet, "SELECT top 1 * FROM con_ordenpagodet", xCon
    '********************************
    RstCab("ano") = AnoTra
    RstCab("idmes") = xMes
    RstCab("fchemi") = CDate(TxtFecha(0).Valor)
    RstCab("fchpag") = CDate(TxtFecha(1).Valor)
    
    RstCab("idprog") = NulosN(lbl_prog(0).Caption)      '--PROGRAMADOR
    RstCab("idest") = NulosN(LblEstado(1).Caption)      '--ESTADO
            
    If opt_operacion(0).Value = True Then RstCab("tipope") = 1
    If opt_operacion(1).Value = True Then RstCab("tipope") = 2
        
    RstCab("idmon") = NulosN(lbl_cb_cod(0).Caption)
    RstCab("numdoc") = Trim(txt(1).Text)
    RstCab("imptot") = NulosN(Trim(txtTotal(2).Text))
    RstCab("impapro") = NulosN(Trim(txtTotal(2).Text)) 'NulosN(Trim(txtApro(2).Text))
    RstCab("imprech") = 0 ' NulosN(Trim(txtRech(2).Text))
        
    RstCab("obs") = Trim(txt(2).Text)

    RstCab.Update
        
    For xFil = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("idord") = xCod
        RstDet("idcom") = NulosN(Fg1.TextMatrix(xFil, 1))
        RstDet("aprobado") = -1
        RstDet("saldo") = NulosN(Fg1.TextMatrix(xFil, 11))
        RstDet("acuenta") = NulosN(Fg1.TextMatrix(xFil, 12))
        RstDet("nuevosaldo") = NulosN(Fg1.TextMatrix(xFil, 13))
        RstDet.Update
    Next xFil
    
    MsgBox "La Programación de Pago se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    xCon.CommitTrans
    Grabar = True
Salir:
    Set RstCab = Nothing
    Set RstDet = Nothing
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + vbCr + Trim(Err.Description), vbCritical, xTitulo
    Grabar = False
    Exit Function
End Function


Private Function fValidarDatos() As Boolean
    If IsDate(TxtFecha(0).Valor) = False Then
        MsgBox "No ha especificado la fecha de emisión", vbInformation, xTitulo
        TxtFecha(0).SetFocus
        Exit Function
    End If
    
    If IsDate(TxtFecha(1).Valor) = False Then
        MsgBox "No ha especificado la fecha de pago", vbInformation, xTitulo
        TxtFecha(1).SetFocus
        Exit Function
    End If
    If CDate(TxtFecha(0).Valor) > CDate(TxtFecha(1).Valor) Then
        MsgBox "La fecha de Pago es inferior a la fecha de emisión" + vbCr + "Modifique la fecha de Pago", vbInformation, xTitulo
        TxtFecha(1).SetFocus
        Exit Function
    End If
    If CDate(TxtFecha(0).Valor) > CDate(TxtFecha(1).Valor) Then
        MsgBox "La fecha de Pago es inferior a la fecha de Emisión" + vbCr + "Modifique la fecha Pago", vbInformation, xTitulo
        TxtFecha(1).SetFocus
        Exit Function
    End If
    
    If lbl_prog(1).Caption = "0" And QueHace = 1 Then
        MsgBox "Ust. No no puede Programar los Pagos", vbInformation, xTitulo
        Exit Function
    End If
    
    Dim band As Integer
    band = Validar(txt_cb)
    If band <> -1 Then
       MsgBox "Llene el Campo de " & lbl_cb_capt(band).Caption, vbInformation, xTitulo
       txt_cb(band).SetFocus
       Exit Function
    End If
    If NulosN(txt_cb(0).Text) = 0 Then
        MsgBox "Llene el Campo de Moneda", vbInformation, xTitulo
        txt_cb(0).SetFocus
        Exit Function
    End If
    band = Validar(txt)
    If band > 0 Then
       MsgBox "Llene el Campo de " & lbl(band).Caption, vbInformation, xTitulo
       txt(band).SetFocus
       Exit Function
    End If

    If Fg1.Rows = 1 Then
        MsgBox "Ingrese por lo menos un Registro de Compra para Programar Pagos", vbExclamation, xTitulo
        Cmd(0).SetFocus
        Exit Function
    End If
    
    '--VALIDAR EL INGRESO DE LOS IMPORTES A PAGAR
    Dim mRow  As Long
    For mRow = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(mRow, 11)) = 0 Then
            MsgBox "Ingrese un valor Acuenta a la Compra:" + vbCr + _
            "Proveedor:        " + Fg1.TextMatrix(mRow, 8) & "" + vbCr + _
            "Num.Reg:         " + Fg1.TextMatrix(mRow, 3) & "" + vbCr + _
            "N°.Documento: " + Fg1.TextMatrix(mRow, 6) & "", vbExclamation, xTitulo
            
            Agregando = True:  Fg1.Row = mRow: Fg1.Col = 11: Agregando = False
            Fg1.SetFocus
            Exit Function
        End If
    Next mRow
    '-----

    fValidarDatos = True
End Function
 

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then Imprimir True

    If ButtonMenu.Index = 2 Then Imprimir
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If QueHace = 3 Then Exit Sub
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

Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)
    
    If txt_cb(Index).Text = "" Then Exit Sub
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
    
        Case 0 '--MONEDA
            nSQL = "SELECT mae_moneda.id, mae_moneda.descripcion,mae_moneda.id as cod " _
                + vbCr + " From mae_moneda " _
                + vbCr + " WHERE (((mae_moneda.id)=" + CStr(Trim(txt_cb(Index).Text)) + "));"
    
    End Select
    If xCon.State = 0 Then Exit Sub
    RST_Busq xRs, nSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount > 0 Then
        txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
        lbl_cb_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cb_cod(Index).Caption = ""
    End If
    Set xRs = Nothing
       
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
    If Index = 1 Then
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub


Private Function fVerificarProgAut(BUSCAPROGRAMADOR As Boolean, OBJ_ID As Label, OBJ_NOMBRE As Label) As Boolean
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQL_PROG As String
    If BUSCAPROGRAMADOR = True Then
        nSQL_PROG = " AND con_emptes.prog=-1; "
    Else
        nSQL_PROG = " AND con_emptes.aut=-1; "
    End If
    'nSQL = "SELECT  con_emptes.id, [pla_empleados].[nom] & ' ' & [pla_empleados].[ape] AS nombre " _
    + vbCr + " FROM (pla_empleados INNER JOIN con_emptes ON pla_empleados.id = con_emptes.idemp) INNER JOIN mae_usuarios ON pla_empleados.id = mae_usuarios.idemp " _
    + vbCr + " WHERE mae_usuarios.id= " + CStr(xIdUsuario) + nSQL_PROG
    nSQL = "SELECT con_emptes.id, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ', ' & [pla_empleados].[nom] AS nombre" _
        & " FROM (pla_empleados INNER JOIN con_emptes ON pla_empleados.id = con_emptes.idemp) INNER JOIN mae_usuarios " _
        & " ON pla_empleados.id = mae_usuarios.idemp WHERE (((mae_usuarios.id)=1) AND ((con_emptes.prog)=-1))"
    
    RST_Busq xRs, nSQL, xCon
    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Then
        OBJ_ID.Caption = "0"
        OBJ_NOMBRE.Caption = "--"
    Else
        OBJ_ID.Caption = xRs.Fields(0) & ""
        OBJ_NOMBRE.Caption = xRs.Fields(1) & ""
        fVerificarProgAut = True
    End If
Salir:
    Set xRs = Nothing
End Function

Private Function fValidarMoneda() As Boolean
    '--FUNCTION QUE VALIDAR SI SELECCIONO LA MONEDA
    If lbl_cb_cod(1).Caption = "" Then
        MsgBox "Seleccione primero la Moneda", vbInformation, xTitulo
        cb_Click 1
        Exit Function
    End If
    fValidarMoneda = True
End Function

'------DEL CAMBIO DE PERIODO
Private Sub CambiarMes()
    
    xMes = SeleccionaMes(xCon)
    pCargarGrid
End Sub

Sub Buscar()
    On Error GoTo error
    TabOne1.CurrTab = 0
     
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
   
    Dim xCampos(8, 4) As String
    
    xCampos(0, 0) = "Tipo Mov.":        xCampos(0, 1) = "operacion":   xCampos(0, 2) = "1000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Fch. Emi":         xCampos(1, 1) = "fchemi1":   xCampos(1, 2) = "850":     xCampos(1, 3) = "F"
    xCampos(2, 0) = "Fch. Pag":         xCampos(2, 1) = "fchpag1":   xCampos(2, 2) = "850":     xCampos(2, 3) = "F"
    xCampos(3, 0) = "M":                xCampos(3, 1) = "simbolo":  xCampos(3, 2) = "450":     xCampos(3, 3) = "C"
    xCampos(4, 0) = "Tot.Imp":          xCampos(4, 1) = "imptot":   xCampos(4, 2) = "1000":     xCampos(4, 3) = "N"
    xCampos(5, 0) = "Tot.Aprob":        xCampos(5, 1) = "impapro":  xCampos(5, 2) = "1000":     xCampos(5, 3) = "N"
    xCampos(6, 0) = "Tot.Recha":        xCampos(6, 1) = "imprech":  xCampos(6, 2) = "1000":     xCampos(6, 3) = "N"
    xCampos(7, 0) = "Emitido Por:":     xCampos(7, 1) = "prog":     xCampos(7, 2) = "2000":    xCampos(7, 3) = "C"
            
    nSQL = "SELECT format(con_ordenpago.fchemi,'dd/m/yy') as fchemi1,format(con_ordenpago.fchpag,'dd/mm/yy') as fchpag1,con_ordenpago.*, pla_empleados.nom & ' ' & pla_empleados.ape AS prog, pla_empleados_1.ape & ' ' & pla_empleados_1.nom AS aut, mae_estados.descripcion AS estdesc, IIf(con_ordenpago.tipope=1,'Caja','Banco') AS operacion , mae_moneda.simbolo " _
        + vbCr + " FROM pla_empleados RIGHT JOIN (mae_moneda RIGHT JOIN (con_emptes RIGHT JOIN (mae_estados RIGHT JOIN ((con_ordenpago LEFT JOIN con_emptes AS con_emptes_1 ON con_ordenpago.idaut = con_emptes_1.id) LEFT JOIN pla_empleados AS pla_empleados_1 ON con_emptes_1.idemp = pla_empleados_1.id) ON mae_estados.id = con_ordenpago.idest) ON con_emptes.id = con_ordenpago.idprog) ON mae_moneda.id = con_ordenpago.idmon) ON pla_empleados.id = con_emptes.idemp " _
        + vbCr + " WHERE con_ordenpago.ano = " & AnoTra & " And con_ordenpago.idmes = " & xMes & " ; "
            
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Programación de Pagos", "fchemi", "prog", Principio
    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir
    
    RstFrm.MoveFirst
    RstFrm.Find "id = " + CStr(xRs("id"))
Salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub




Private Sub Filtrar()
    
    Dim xCampos(6, 4) As String
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    xCampos(0, 0) = "Tipo Mov.":        xCampos(0, 1) = "tipmov":   xCampos(0, 2) = "C":     xCampos(0, 3) = "1000"
    xCampos(1, 0) = "Fch. Emi":         xCampos(1, 1) = "fchemi":   xCampos(1, 2) = "F":     xCampos(1, 3) = "850"
    xCampos(2, 0) = "M":                xCampos(2, 1) = "simbolo":  xCampos(2, 2) = "C":     xCampos(2, 3) = "450"
    xCampos(3, 0) = "Importe":          xCampos(3, 1) = "imp":      xCampos(3, 2) = "N":     xCampos(3, 3) = "800"
    xCampos(4, 0) = "Emitido Por:":     xCampos(4, 1) = "prog":     xCampos(4, 2) = "C":     xCampos(4, 3) = "2500"
    xCampos(5, 0) = "T. Persona":       xCampos(5, 1) = "tipper":   xCampos(5, 2) = "C":     xCampos(5, 3) = "1000"
    xCampos(6, 0) = "Entregado A":      xCampos(6, 1) = "benef":    xCampos(6, 2) = "C":     xCampos(6, 3) = "2500"
    
    
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg3

    TabOne1.CurrTab = 0
End Sub


Private Sub Imprimir(Optional IMP_LISTADO As Boolean = False)

    On Error GoTo error

    Me.MousePointer = vbHourglass
    If IMP_LISTADO = False Then
        If Me.TabOne1.CurrTab = 0 Then
        
        Else
''            MsgBox "Primero muestre el detalle del Registro" + vbCr + _
''                   "Luego inténtelo otra vez", vbExclamation, xTitulo
        End If
    Else
    
        TDB_IMPRIMIR Dg3, "IMPRESIÓN DE CUENTAS POR RENDIR", "LISTADO DE CUENTAS POR RENDIR-  Periodo: " + MonthName(xMes, False)
   
    End If

    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "IMPRIMIR"

End Sub


Private Sub CmdModificar_Click(Index As Integer)
    Select Case Index
        Case 0 '--modificar
            If CmdModificar(0).Caption = "&Modificar" Then
                CmdModificar(0).Caption = "&Grabar"
                pBloqueaModificar False
                QueHace = 2
                 Fg1.SelectionMode = flexSelectionFree
                TxtFecha(1).SetFocus
            Else
                '---
                If fGrabarMofificar() = True Then
                    CmdModificar_Click 1
                End If
                '---
            End If
        Case 1 '--cancelar
            CmdModificar(0).Caption = "&Modificar"
            pBloqueaModificar True
             Fg1.SelectionMode = flexSelectionByRow
            QueHace = 3
    End Select
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
If Index = 1 Then
    If Trim(txt(1).Text) <> "" Then
        txt(1).Text = Format(txt(1).Text, "0000000000")
    End If
End If
End Sub

Private Sub txtfecha_Validate(Index As Integer, Cancel As Boolean)
    If Index <> 0 Then Exit Sub
    If IsDate(TxtFecha(0).Valor) = True Then
        LblTipoCambio.Caption = HallaTipoCambio(TxtFecha(0).Valor, 2, Venta, xCon)
    Else
        LblTipoCambio.Caption = ""
    End If
End Sub


Private Sub pGenerarAsiento(RstDiario As ADODB.Recordset, nAnoTrabajo, mMesActivo, IDLibro, IDMov, mIdDocPro, mCorr, nAsiento, mTipoCambio, FchDoc, IDcuenta, IDMoneda, mImporte, Optional EsDEBE As Boolean)
    '--mCorr por le general es igual a 0
    RstDiario.AddNew
    RstDiario("año") = nAnoTrabajo
    RstDiario("idmes") = mMesActivo  'CODIGO DEL MES
    RstDiario("idlib") = IDLibro     'CODIGO DEL LIBRO
    RstDiario("idmov") = IDMov       'CODIGO DEL MOVIMIENTO
    RstDiario("iddocpro") = mIdDocPro
    RstDiario("correlativo") = mCorr
    RstDiario("numasi") = nAsiento
    RstDiario("tc") = mTipoCambio
    If mMesActivo = 0 Then
        RstDiario("fchasi") = CDate("01/01/" + nAnoTrabajo)
    ElseIf mMesActivo = 13 Then
        RstDiario("fchasi") = CDate("31/12/" + nAnoTrabajo)
    Else
        RstDiario("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + nAnoTrabajo)
    End If
    RstDiario("fchdoc") = FchDoc
    RstDiario("idcue") = IDcuenta
    If EsDEBE = False Then
        If IDMoneda = 1 Then '--DE LA MONEDA
            RstDiario("imphabsol") = mImporte
            RstDiario("imphabdol") = 0
        Else
            RstDiario("imphabsol") = mImporte * mTipoCambio
            RstDiario("imphabdol") = mImporte
        End If
    Else
        If IDMoneda = 1 Then '--DE LA MONEDA
            RstDiario("impdebsol") = mImporte
            RstDiario("impdebdol") = 0
        Else
            RstDiario("impdebsol") = mImporte * mTipoCambio
            RstDiario("impdebdol") = mImporte
        End If
    End If

    RstDiario.Update
End Sub

Function fGrabarMofificar() As Boolean
    If IsDate(TxtFecha(0).Valor) = False Then
        MsgBox "Falta ingresar la Fecha a Emisión", vbExclamation, xTitulo
        TxtFecha(0).SetFocus
        Exit Function
    End If
    If IsDate(TxtFecha(1).Valor) = False Then
        MsgBox "Falta ingresar la Fecha de Pago ", vbExclamation, xTitulo
        TxtFecha(1).SetFocus
        Exit Function
    End If
    If CDate(TxtFecha(0).Valor) > CDate(TxtFecha(1).Valor) Then
        MsgBox "La Fecha de Pago es inferior a la Fecha de Emisión " + vbCr + "Modifique la Fecha de Pago", vbExclamation, xTitulo
        TxtFecha(1).SetFocus
        Exit Function
    End If
    
    If NulosN(txt_cb(0).Text) = 0 Then
        MsgBox "Falta ingresar la Moneda", vbExclamation, xTitulo
        txt_cb(0).SetFocus
        Exit Function
    End If
    
    If MsgBox("Seguro desea modificar el registro", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo Salir
    
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xCod As Integer
    Dim xFil&
    
    On Error GoTo LaCague

    xCon.BeginTrans
    
    If QueHace = 1 Then
        Exit Function
    Else
        xCod = RstFrm("id")
        RST_Busq RstCab, "SELECT * FROM con_ordenpago WHERE id =" & xCod & "", xCon
        '--ELIMINANDO LOS REGISTROS DEL DETALLE DE ORDEN DE PAGO
        xCon.Execute "Delete From con_ordenpagodet Where idord= " & xCod & ";"
    End If
    '********************************
    RST_Busq RstDet, "SELECT top 1 * FROM con_ordenpagodet", xCon
    '********************************
    RstCab("fchpag") = CDate(TxtFecha(1).Valor)
    RstCab("imptot") = NulosN(Trim(txtTotal(2).Text))
    RstCab("impapro") = NulosN(Trim(txtApro(2).Text))
    RstCab("imprech") = NulosN(Trim(txtRech(2).Text))
    RstCab("obs") = Trim(txt(2).Text)

    RstCab.Update
        
    For xFil = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("idord") = xCod
        RstDet("idcom") = NulosN(Fg1.TextMatrix(xFil, 1))
        RstDet("aprobado") = NulosN(Fg1.TextMatrix(xFil, 2))
        RstDet("saldo") = NulosN(Fg1.TextMatrix(xFil, 11))
        RstDet("acuenta") = NulosN(Fg1.TextMatrix(xFil, 12))
        RstDet("nuevosaldo") = NulosN(Fg1.TextMatrix(xFil, 13))
        RstDet.Update
    Next xFil
    
    MsgBox "La Prgramación de Pago se modificó con éxito", vbInformation, xTitulo
        
    xCon.CommitTrans
    fGrabarMofificar = True
    RstFrm.Requery
    If RstFrm.RecordCount <> 0 Then
        RstFrm.MoveFirst
        RstFrm.Find "id= " & xCod
    End If
    
Salir:
    Set RstCab = Nothing
    Set RstDet = Nothing
    Exit Function
    
LaCague:
    Resume
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo modificar el registro por el siguiente motivo :" + Trim(Err.Description), vbCritical, xTitulo
    Exit Function
End Function


Private Sub pBloqueaModificar(band As Boolean)

    TxtFecha(0).Locked = band
    txt(2).Locked = band
    habilitar cmd_estado, band
    CmdModificar(1).Enabled = Not band
    
    habilitar cb, band
    habilitar_Locked txt_cb, Not band
    habilitar Cmd, Not band
    ChkAutorizar.Enabled = Not band
End Sub


Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 0 '--AGREGAR COMPRAS YA REGISTRADAS
            pRegistroAdd False
        Case 1 '--SELECCCIONAR
            pRegistroAdd True
        Case 2 '--ELIMINAR REGISTROS AGREGADOS
            pRegistroDel
            
    End Select
End Sub

Private Sub pRegistroAdd(Optional fSeleccionVarios As Boolean = True)
    '--CARGAR LAS COMPRAS PARA LUEGO SELECCIONAR LOS QUE DESEEMOS
    '--SE CARGARAN DE ACUERDO A LA MONEDA DE CUENTAS POR RENDIR
    Dim nSQLIdCompra As String
    
    If NulosN(lbl_cb_cod(0).Caption) = 0 Then
        MsgBox "Falta especificar la moneda", vbExclamation, xTitulo
        txt_cb(0).SetFocus
        Exit Sub
    End If
    
    '--GENERAR EL WHERE DE LOS ID'S COMPRA PARA QUE NO SE REPITAN
    nSQLIdCompra = GENERAR_SQL_ID(Fg1, 1, "com_compras.id", " NOT IN ")
    If nSQLIdCompra <> "" Then nSQLIdCompra = " AND " + nSQLIdCompra
    '--DE LA MONEDA
    nSQLIdCompra = nSQLIdCompra + " AND com_compras.idmon=" & NulosN(lbl_cb_cod(0).Caption) & " "
    '----
    On Error GoTo error
    Dim xRs  As New ADODB.Recordset
    Dim xCampos(9, 5) As String
    Dim nSQL As String
    
    xCampos(0, 0) = "Num.Reg.":      xCampos(0, 1) = "registro":      xCampos(0, 2) = "900":    xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
    xCampos(1, 0) = "T.D.":          xCampos(1, 1) = "abrev":        xCampos(1, 2) = "450":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
    xCampos(2, 0) = "M":             xCampos(2, 1) = "simbolo":     xCampos(2, 2) = "450":      xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
    xCampos(3, 0) = "N°.Documento":  xCampos(3, 1) = "doc":         xCampos(3, 2) = "1400":     xCampos(3, 3) = "C":    xCampos(3, 4) = "N"
    xCampos(4, 0) = "Fch.Emi.":      xCampos(4, 1) = "fchdoc":      xCampos(4, 2) = "900":      xCampos(4, 3) = "C":    xCampos(4, 4) = "N"
    xCampos(5, 0) = "Fch.Venc.":     xCampos(5, 1) = "fchven":      xCampos(5, 2) = "900":     xCampos(5, 3) = "C":    xCampos(5, 4) = "N"
    If fSeleccionVarios = False Then
        xCampos(6, 0) = "Proveedor":     xCampos(6, 1) = "nombre":      xCampos(6, 2) = "1500":     xCampos(6, 3) = "C":    xCampos(6, 4) = "N"
    Else
        xCampos(6, 0) = "Proveedor":     xCampos(6, 1) = "nombre":      xCampos(6, 2) = "3600":     xCampos(6, 3) = "C":    xCampos(6, 4) = "N"
    End If
    xCampos(7, 0) = "Importe":       xCampos(7, 1) = "imptot":      xCampos(7, 2) = "800":      xCampos(7, 3) = "N":    xCampos(7, 4) = "N"
    xCampos(8, 0) = "Saldo":         xCampos(8, 1) = "impsal":      xCampos(8, 2) = "800":      xCampos(8, 3) = "N":    xCampos(8, 4) = "N"
    '--obtenemos la consulta
    nSQL = fGenerarConsulta(True, -1, nSQLIdCompra)
    
    If fSeleccionVarios = True Then
        CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), "Buscando Compras"
    Else
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Compras", "doc", "doc", CualquierParte
    End If
    
    Agregando = True
    If xRs.State = 0 Then GoTo Salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo Salir
    If fSeleccionVarios = True Then xRs.MoveFirst
    Do While Not xRs.EOF
        With Fg1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = xRs.Fields("id") & ""
            .TextMatrix(.Rows - 1, 3) = xRs.Fields("registro") & ""
            .TextMatrix(.Rows - 1, 4) = xRs.Fields("abrev") & ""
            .TextMatrix(.Rows - 1, 5) = xRs.Fields("simbolo") & ""
            .TextMatrix(.Rows - 1, 6) = xRs.Fields("doc") & ""
            .TextMatrix(.Rows - 1, 7) = xRs.Fields("fchdoc") & ""
            .TextMatrix(.Rows - 1, 8) = xRs.Fields("fchven") & ""
            .TextMatrix(.Rows - 1, 9) = xRs.Fields("nombre") & ""
            .TextMatrix(.Rows - 1, 10) = Format(NulosN(xRs.Fields("imptot")), FORMAT_MONTO)
            .TextMatrix(.Rows - 1, 11) = Format(NulosN(xRs.Fields("impsal")), FORMAT_MONTO)
            
            
            '---
        End With
        If fSeleccionVarios = False Then Exit Do
        xRs.MoveNext
    Loop
    pTotalizarDatos
    Fg1.Col = 11
    Fg1.Row = 1
    Fg1.SetFocus
    
Salir:
    Agregando = False
    Set xRs = Nothing
    '----
        
    Exit Sub
error:
    Agregando = False
    Set xRs = Nothing
    SHOW_ERROR
End Sub

Private Sub pRegistroDel()
    If Fg1.Row <= 0 Then
        MsgBox "La fila no se puede eliminar" + vbCr + "Seleccione una correcta", vbExclamation, xTitulo
        Exit Sub
    End If
    If MsgBox("Seguro desea Eliminar el registro seleccionado", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    '--ELIMINAR EL REGISTRO
    Fg1.RemoveItem (Fg1.Row)
    If Fg1.Rows > 1 Then Fg1.Row = 1
    pTotalizarDatos
End Sub

Private Function pTotalizarDatos()
    '--
    Dim s&
    '--OBTENER LOS TOTALES
    txtTotal(0).Text = Format(GRID_SUMAR_COL(Fg1, 10), FORMAT_MONTO)
    txtTotal(1).Text = Format(GRID_SUMAR_COL(Fg1, 11), FORMAT_MONTO)
    txtTotal(2).Text = Format(GRID_SUMAR_COL(Fg1, 12), FORMAT_MONTO)
    txtTotal(3).Text = Format(GRID_SUMAR_COL(Fg1, 13), FORMAT_MONTO)
    
    txtApro(0).Text = "0.00"
    txtApro(1).Text = "0.00"
    txtApro(2).Text = "0.00"
    txtApro(3).Text = "0.00"
            
    '--OBTENER SOLO LOS APROBADOS
    For s = 1 To Fg1.Rows - 1
        If Abs(NulosN(Fg1.TextMatrix(s, 2))) = 1 Then
            txtApro(0).Text = Format(NulosN(txtApro(0).Text) + NulosN(Fg1.TextMatrix(s, 10)), FORMAT_MONTO)
            txtApro(1).Text = Format(NulosN(txtApro(1).Text) + NulosN(Fg1.TextMatrix(s, 11)), FORMAT_MONTO)
            txtApro(2).Text = Format(NulosN(txtApro(2).Text) + NulosN(Fg1.TextMatrix(s, 12)), FORMAT_MONTO)
            txtApro(3).Text = Format(NulosN(txtApro(3).Text) + NulosN(Fg1.TextMatrix(s, 13)), FORMAT_MONTO)
        End If
    Next s
            
    '--OBTENER SOLO LOS DESAPROBADOS
    txtRech(0).Text = Format(NulosN(txtTotal(0).Text) - NulosN(txtApro(0).Text), FORMAT_MONTO)
    txtRech(1).Text = Format(NulosN(txtTotal(1).Text) - NulosN(txtApro(1).Text), FORMAT_MONTO)
    txtRech(2).Text = Format(NulosN(txtTotal(2).Text) - NulosN(txtApro(2).Text), FORMAT_MONTO)
    txtRech(3).Text = Format(NulosN(txtTotal(3).Text) - NulosN(txtApro(3).Text), FORMAT_MONTO)

End Function


Private Function fGenerarConsulta(fAddRegistro As Boolean, Optional mIdCompra As Integer = -1, Optional nSQLNotIn As String = "") As String
    '--mIdCompra <>-1 CUANDO SE CREA EL REGISTRO DE COMPRA
    Dim nSQL As String
    Dim nnSQLIdCompra As String
    If mIdCompra <> -1 Then nnSQLIdCompra = " AND com_compras.ID = " + CStr(mIdCompra) + " "
    
    If fAddRegistro = True Then '--NUEVO
        nSQL = "SELECT com_compras.id, Left([numreg],2) & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & Right([numreg],4) AS registro, mae_documento.abrev, mae_moneda.simbolo, com_compras!numser & ' ' & com_compras!numdoc AS doc, Format(com_compras.fchdoc,'dd/mm/yy') AS fchdoc, format(com_compras.fchven,'dd/mm/yy') as fchven, mae_prov.nombre, com_compras.imptot, com_compras.impsal " _
            + vbCr + " FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro " _
            + vbCr + " WHERE com_compras.impsal <> 0 " + nnSQLIdCompra + nSQLNotIn _
            + vbCr + " ORDER BY com_compras.fchven asc;"

    Else '--CONSULTA O MODIFICAR
        nSQL = " SELECT com_compras.id as idcom, Left([numreg],2) & IIf(mae_libros.codsun Is Null Or mae_libros.codsun='','FF',mae_libros.codsun) & Right([numreg],4) AS registro, mae_documento.abrev, mae_moneda.simbolo, com_compras!numser & ' ' & com_compras!numdoc AS doc, Format(com_compras.fchdoc,'dd/mm/yy') AS fchdoc, format(com_compras.fchven,'dd/mm/yy') as fchven, mae_prov.nombre, com_compras.imptot, con_ordenpagodet.saldo, con_ordenpagodet.acuenta, con_ordenpagodet.nuevosaldo, con_ordenpagodet.aprobado " _
            + vbCr + " FROM (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro) INNER JOIN con_ordenpagodet ON com_compras.id = con_ordenpagodet.idcom " _
            + vbCr + " WHERE con_ordenpagodet.idord = " & NulosN(RstFrm.Fields("id")) & "" _
            + vbCr + " ORDER BY  com_compras.fchven asc;"

        
    End If
    
    fGenerarConsulta = nSQL
End Function


Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Or KeyCode = vbKeyInsert Then 'F3 = Agregar Item
        Cmd_Click 1
    End If
    If KeyCode = 115 Or KeyCode = vbKeyDelete Then
        Cmd_Click 2  'F4 = Eliminar Item
    End If
End Sub
Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    If Agregando = True Then Exit Sub
    If QueHace = 3 Then Exit Sub
    If Row = 0 Then Exit Sub
    Select Case Col
        Case 12
            If IsNumeric(Fg1.TextMatrix(Row, Col)) = False Then
                MsgBox "El valor ingresado no es numérico", vbCritical, xTitulo
                Fg1.TextMatrix(Row, 12) = "":       Fg1.TextMatrix(Row, 13) = ""
            Else
                If NulosN(Fg1.TextMatrix(Row, 12)) > NulosN(Fg1.TextMatrix(Row, 11)) Then
                    MsgBox "El valor Ingresado supera al saldo anterior" + vbCr + "Saldo Anterior: " & NulosN(Fg1.TextMatrix(Row, 11)), vbExclamation, xTitulo
                    Fg1.TextMatrix(Row, 12) = "":        Fg1.TextMatrix(Row, 13) = ""
                    Exit Sub
                End If
                Fg1.TextMatrix(Row, 13) = NulosN(Fg1.TextMatrix(Row, 11)) - NulosN(Fg1.TextMatrix(Row, 12))
            End If
    End Select
    pTotalizarDatos
    Exit Sub
error:
    SHOW_ERROR
End Sub


Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace <> 3 Then
            PopupMenu Menu1
        End If
    End If
End Sub


Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    If Fg1.Col = 2 Or Fg1.Col = 12 Then
        If Fg1.Col = 2 Then
            If CmdModificar(0).Visible = True Then   '--si grabar esta dehabilitado
                Fg1.Editable = flexEDKbdMouse
            Else
                Fg1.Editable = flexEDNone
            End If
        Else
            Fg1.Editable = flexEDKbdMouse
        End If
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub
Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case Col
        Case 2
        Case 12
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub


