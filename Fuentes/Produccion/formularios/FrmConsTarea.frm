VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsTarea 
   Caption         =   "Producción - Consulta de Tareas"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   615
   ClientWidth     =   11820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11820
   Begin VB.Frame FraOpcion 
      BorderStyle     =   0  'None
      Height          =   2550
      Left            =   3660
      TabIndex        =   33
      Top             =   2640
      Visible         =   0   'False
      Width           =   4020
      Begin VB.TextBox txtLote 
         Height          =   285
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   47
         Text            =   "txtLote"
         Top             =   1065
         Width           =   1605
      End
      Begin VB.CheckBox chk 
         Caption         =   "Inconsistencias"
         Height          =   195
         Index           =   2
         Left            =   2280
         TabIndex        =   46
         Top             =   330
         Width           =   1500
      End
      Begin VB.Frame Frame7 
         Caption         =   "Seleccionar Tipo de Tarea"
         Height          =   600
         Left            =   120
         TabIndex        =   41
         Top             =   1380
         Width           =   3765
         Begin VB.OptionButton OptTarea 
            Caption         =   "Indirecto"
            Height          =   195
            Index           =   2
            Left            =   2580
            TabIndex        =   44
            Top             =   270
            Width           =   1005
         End
         Begin VB.OptionButton OptTarea 
            Caption         =   "Directo"
            Height          =   195
            Index           =   1
            Left            =   1350
            TabIndex        =   43
            Top             =   300
            Width           =   885
         End
         Begin VB.OptionButton OptTarea 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   42
            Top             =   300
            Value           =   -1  'True
            Width           =   885
         End
      End
      Begin VB.TextBox txtObs 
         Height          =   285
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   40
         Text            =   "txtObs"
         Top             =   750
         Width           =   1605
      End
      Begin VB.CheckBox chk 
         Caption         =   "Mostrar sólo Reproceso"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   38
         Top             =   555
         Width           =   2070
      End
      Begin VB.CheckBox chk 
         Caption         =   "Mostrar sólo Observados"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   330
         Width           =   2070
      End
      Begin VB.CommandButton CmdEditor 
         Caption         =   "Aceptar"
         Height          =   465
         Index           =   3
         Left            =   1275
         TabIndex        =   35
         Top             =   2010
         Width           =   1365
      End
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   3735
         Picture         =   "FrmConsTarea.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   34
         ToolTipText     =   "Cerrar"
         Top             =   30
         Width           =   195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Filtrar por Nº Lote"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   1155
         Width           =   1230
      End
      Begin VB.Label LblObs 
         AutoSize        =   -1  'True
         Caption         =   "Filtrar de Observaciones"
         Height          =   195
         Left            =   150
         TabIndex        =   39
         Top             =   840
         Width           =   1710
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   1
         X1              =   0
         X2              =   15
         Y1              =   15
         Y2              =   5070
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   6435
         Y1              =   2535
         Y2              =   2535
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -150
         X2              =   5895
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   4005
         X2              =   4005
         Y1              =   -135
         Y2              =   4755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Opciones"
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
         TabIndex        =   36
         Top             =   60
         Width           =   810
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   270
         Index           =   0
         Left            =   30
         Top             =   15
         Width           =   3930
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4860
         Top             =   90
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
               Picture         =   "FrmConsTarea.frx":02EC
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsTarea.frx":0830
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsTarea.frx":0BC2
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsTarea.frx":0D46
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsTarea.frx":119A
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsTarea.frx":12B2
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsTarea.frx":17F6
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsTarea.frx":1D3A
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsTarea.frx":1E4E
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsTarea.frx":1F62
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsTarea.frx":23B6
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsTarea.frx":2522
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsTarea.frx":2A6A
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsTarea.frx":2D84
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraProgreso 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   2835
      TabIndex        =   12
      Top             =   3720
      Visible         =   0   'False
      Width           =   5760
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   90
         TabIndex        =   13
         Top             =   345
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tareas"
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
         Index           =   1
         Left            =   1185
         TabIndex        =   16
         Top             =   75
         Width           =   585
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   90
         TabIndex        =   15
         Top             =   75
         Width           =   1035
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Index           =   2
         Left            =   4140
         TabIndex        =   14
         Top             =   75
         Width           =   1530
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5745
         X2              =   5745
         Y1              =   -90
         Y2              =   4800
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -150
         X2              =   5895
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   2
         X1              =   -60
         X2              =   6360
         Y1              =   690
         Y2              =   690
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
   End
   Begin VB.Frame Frame2 
      Height          =   2280
      Left            =   0
      TabIndex        =   1
      Top             =   255
      Width           =   11850
      Begin VB.CommandButton CmdOpcion 
         Caption         =   "Opciones"
         Height          =   375
         Left            =   9720
         TabIndex        =   45
         Top             =   180
         Width           =   1905
      End
      Begin VB.Frame Frame3 
         Caption         =   "Seleccionar Area"
         Height          =   495
         Left            =   5730
         TabIndex        =   27
         Top             =   90
         Width           =   3885
         Begin VB.CommandButton cb 
            Height          =   240
            Index           =   0
            Left            =   1005
            Picture         =   "FrmConsTarea.frx":3116
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   210
            Width           =   225
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   0
            Left            =   630
            MaxLength       =   12
            TabIndex        =   29
            Text            =   "txt_cb(0)"
            Top             =   180
            Width           =   615
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Area"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   32
            Top             =   270
            Width           =   330
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
            Left            =   2610
            TabIndex        =   31
            Top             =   180
            Visible         =   0   'False
            Width           =   1005
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
            Left            =   1275
            TabIndex        =   30
            Top             =   180
            Width           =   2475
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Consulta"
         Height          =   855
         Left            =   2085
         TabIndex        =   22
         Top             =   90
         Width           =   2145
         Begin VB.OptionButton OptConsulta 
            Caption         =   "Eficiencia del Personal"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   427
            Width           =   1905
         End
         Begin VB.OptionButton OptConsulta 
            Caption         =   "Control de Tarea"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   225
            Value           =   -1  'True
            Width           =   1725
         End
         Begin VB.OptionButton OptConsulta 
            Caption         =   "Mínimos y Máximos"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   615
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Agrupar Por"
         Height          =   855
         Left            =   4275
         TabIndex        =   5
         Top             =   90
         Width           =   1380
         Begin VB.OptionButton OptGrupo 
            Caption         =   "x Tarea"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   17
            Top             =   615
            Width           =   855
         End
         Begin VB.OptionButton OptGrupo 
            Caption         =   "x &Producto"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   225
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton OptGrupo 
            Caption         =   "x Personal"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   2
            Top             =   427
            Width           =   1215
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   1245
         Index           =   0
         Left            =   45
         TabIndex        =   3
         ToolTipText     =   "Buscar Personal"
         Top             =   1020
         Width           =   3855
         _cx             =   6800
         _cy             =   2196
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsTarea.frx":3248
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
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   1245
         Index           =   1
         Left            =   7905
         TabIndex        =   4
         ToolTipText     =   "Buscar Productos"
         Top             =   1020
         Width           =   3855
         _cx             =   6800
         _cy             =   2196
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
         AllowUserResizing=   0
         SelectionMode   =   1
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
         FormatString    =   $"FrmConsTarea.frx":32A6
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
         PicturesOver    =   -1  'True
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
      Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
         Height          =   300
         Index           =   0
         Left            =   645
         TabIndex        =   7
         Top             =   240
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
         Left            =   645
         TabIndex        =   8
         Top             =   615
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
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   1245
         Index           =   2
         Left            =   3975
         TabIndex        =   11
         ToolTipText     =   "Buscar Tareas"
         Top             =   1020
         Width           =   3855
         _cx             =   6800
         _cy             =   2196
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
         AllowUserResizing=   0
         SelectionMode   =   1
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
         FormatString    =   $"FrmConsTarea.frx":331B
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
         PicturesOver    =   -1  'True
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
      Begin VB.Frame Fra_Top 
         Enabled         =   0   'False
         Height          =   450
         Left            =   5730
         TabIndex        =   18
         Top             =   495
         Width           =   3885
         Begin VB.TextBox txt_top 
            Height          =   315
            Left            =   630
            MaxLength       =   2
            TabIndex        =   20
            Text            =   "txt_top"
            Top             =   120
            Width           =   615
         End
         Begin VB.ComboBox cb_top 
            Height          =   315
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label lbl_top 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mostrar"
            Height          =   195
            Left            =   60
            TabIndex        =   21
            Top             =   225
            Width           =   525
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   705
         Width           =   465
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   5025
      Left            =   0
      TabIndex        =   26
      Top             =   2580
      Width           =   11850
      _cx             =   20902
      _cy             =   8864
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
      FrontTabColor   =   -2147483644
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "     Detalle     |     Resumen     "
      Align           =   0
      CurrTab         =   0
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
         Height          =   4605
         Left            =   45
         TabIndex        =   49
         Top             =   45
         Width           =   11760
         _cx             =   20743
         _cy             =   8123
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsTarea.frx":338D
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
         Height          =   4605
         Left            =   12495
         TabIndex        =   50
         Top             =   45
         Width           =   11760
         _cx             =   20743
         _cy             =   8123
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsTarea.frx":359E
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
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu Menu2_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu2_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu2_3 
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "FrmConsTarea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMCONSTAREA.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO QUE MUESTRA LAS TAREAS REGISTRADAS
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 29/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim BAND_INTERRUMPIR As Boolean     ' SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA TRUE SE INTERRUMPE
' DE LA IMPRESION
Dim T_RPT_PERIODO As String         ' PERIODO DEL REPORTE
Dim T_RPT_TITULO As String          ' TITULO DE REPORTE
Dim ARR_ANYO() As String            ' ARRAY DE AÑOS SELECCIONADOS
Dim ARR_XX() As String              ' SE CARGARA CUANDO SE CARGA EL FORMULARIO Y CUANDO SE CAMBIE EL ESTILO(MES, TRIMESTRE,SEMESTRE)
Dim ARR_TMP(3, 1) As String         ' 0 = PROGRAMADO=>> 0::TOTAL,1::TOTAL GEN
                                    ' 1 = TEORICO=>> 0::TOTAL,1::TOTAL GEN
                                    ' 2 = REAL=>> 0::TOTAL,1::TOTAL GEN
                                    ' 3 = DIF=>> 0::TOTAL,1::TOTAL GEN
Dim Q_TOTAL_ANYO As Integer         ' INDICA LA CANTIDAD DE AÑOS DE BUSQUEDA,
                                    ' EJ. 2004,2005 => Q_TOTAL_ANYO = 2
                                    ' EJ. 2004,2005,2006 => Q_TOTAL_ANYO = 3
Dim Q_COL_FILA As Integer           ' INDICA LA CANTIDAD DE COLUMNAS ANTES DE LOS MESES
                                    ' EJ. IDCLI,CLIENTE => Q_COL_FILA=2
                                    'IDCLI,ID_PTO_VTA,CLIENTE,PTO_VENTA => Q_COL_FILA=4
Dim Q_COL_FILA_ULTIMO As Integer    ' INDICA LA CANTIDAD DE COLUMNAS ADICIONALES QUE SE COLOCARAN DESPUES DEL TOTAL
Dim Q_POS_MES_INICIO As Integer     ' INDICA LA POSICION INICIAL DE LA COLUMNA DEL PRIMER MES, NO CAMBIA
                                    ' EJ. Q_POS_MES_INICIO = Q_COL_FILA +1
Dim Q_POS_MES As Integer            ' INDICA LA POSICION DEL MES, ESTO CAMBIA
                                    ' UTIL PARA COLOCAR LOS DATOS EN EL GRID
Dim Q_COL_FILA_OCULTA As Integer    ' INDICA LAS COLUMNAS QUE CONTENDRAN LOS ID'S, ESTOS SE OCULTARAN
                                    ' -1 NO SE OCULTA, <> -1 SE PROCEDE A ACULTAR
                                    ' EJ. CLIENTE  vta_ventas.idcli,
                                    ' PUNTO DE VENTA vta_guia.idpunven
                                    ' PRODUCTO   alm_inventario.tippro
                                    ' ITEM       alm_inventario.id
                                    ' EMPLEADO   vta_ventas.idven
Dim Q_POSICION_TOTAL  As Integer    ' INDICA LA POCISION DE LA COLUMNA DONDE SE COLOCARA EL NOMBRE DEL TOTAL Y TOTAL_GRL
                                    ' OBTENDRA VALOR EN fGenerarConsulta()
Dim Q_COL_COMPARAR_GRUPO As Integer ' INDICA LA COLUMNA PARA COMPARAR EL GRUPO
                                    ' OBTENDRA VALOR EN fGenerarConsulta()
Dim Q_COL_GRUPO_ADD As Integer      ' ADICIONAR DATOS AL GRID EN EL GRUPO (EJ. Q_COL_GRUPO_ADD=2 =>> NOMBRE_GRUPO|COLUM1|COLUM2)
                                    ' FNUCIONA SI Q_COL_GRUPO_ADD<>-1
Dim N_CAMPO_GRUPO_ADD As String     ' INDICA EL NOMBRE DEL CAMPO A COMPARAR PARA AGREGAR AL LA FILA DEL GRUPO DEPENDE DE Q_COL_GRUPO_ADD
Dim Q_COL_GRUPO_INICIO As Integer   ' INDICA EL INICIO DE LA COLUMNA DEL GRUPO,
Dim Q_COL_GRUPO_TERMINA As Integer  ' INDICA EL TERMINO DE LA COLUMNA DEL GRUPO, UNE LAS CELDAS DE [Q_COL_GRUPO_INICIO] HASTA [Q_COL_GRUPO_TERMINA]
Dim Q_COL_ARR_TOTAL As Integer      ' NOS INDICA EL TOTAL DE COLUMNAS QUE VA A CONTENER LOS ACUMULADOS
                                    ' OBTENDRA VALOR EN fValidarConsulta()
                                    ' SI SEL MES: ENE, FEB, MAR => Q_COL_ARR_TOTAL= 2
                                    ' SI SEL TRI: ENE-MAR, ABR-JUN => Q_COL_ARR_TOTAL= 1 OBS: SE INICIA DESDE POS=0
Dim F_ES_COMPRA As Boolean          ' INDICA SI ES COMPRA O VENTA
                                    ' TRUE::ES COMPRA, FALSE::ES VENTA
Dim ID_PROGRAMA As String
Dim ID_RECETA As String
Dim TIPO_VENTANA As e_PROGRAMA
Dim ESTILO_VISTA As Integer
Dim nSQLValor_FONDO As String       ' AMACENA EL VALOR PARA COMPARAR
Dim nSQLValor_FONDO_COLOR As Long   ' AMACENA EL VALOR DEL COLOR PARA EL FONDO DE LA FILA
Dim F_CAMIAR_FONDO As Boolean       ' FALSE::SE CONSERVA EL FONDO ACTUAL, TRUE::CAMBIA DE FONDO
Dim Q_COL_COMPARAR_FONDO As Integer ' INDICA LA COLUMNA DEL RECORDSET QUE DEBERA DE COMPARAR PARA CAMBIAR DE FONDO -1=NO HACER NADA
Dim SeEjecuto  As Boolean
'------------

'*****************************************************************************************************
'* Nombre           : pConsultar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRAS LAS TAREAS REALIZADAS EN EL PERIODO ESPECIFICADO POR EL USUARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pConsultar()
    Dim rst_select As New ADODB.Recordset
    Dim nSQLSelect As String            ' RECIBIR LA CONSULTA
    Dim mTipoConsulta As Integer        ' valor para configurar el tipo de consulta y obtener el script sql
        
    If fValidarConsulta() = False Then Exit Sub
    
    BAND_INTERRUMPIR = False
    ' CONFIGURAR LA PRESENTACION DE LA CONSULTA
    LimpiarGrid Me.Fg1, False, 1
    LimpiarGrid Me.Fg2, False, 1
    
    ' procesar el detalle
    mTipoConsulta = fEstiloConsulta(1)
    nSQLSelect = fGenerarConsulta(mTipoConsulta)
    pConfigurarGrilla Fg1, mTipoConsulta
    Me.MousePointer = vbHourglass
    DoEvents
    nSQLValor_FONDO = ""
    If nSQLSelect <> "" Then
        PosicionarProgBar
        DoEvents
        ' cargando el detalle
        RST_Busq rst_select, nSQLSelect, xCon
        pCargarDatosGrid Fg1, rst_select
    End If
    ' procesar el resumen
    nSQLValor_FONDO = ""
    mTipoConsulta = fEstiloConsulta(2)
    If mTipoConsulta < 12 Or mTipoConsulta > 14 Then
        nSQLSelect = fGenerarConsulta(mTipoConsulta)
        pConfigurarGrilla Fg2, mTipoConsulta
        If nSQLSelect <> "" Then
             PosicionarProgBar
             DoEvents
             ' cargando el resumen
             Set rst_select = Nothing
             RST_Busq rst_select, nSQLSelect, xCon
             pCargarDatosGrid Fg2, rst_select
        End If
    End If

SALIR:
    FraProgreso.Visible = False
    Set rst_select = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    FraProgreso.Visible = False
    Set rst_select = Nothing
    SHOW_ERROR Me.Name, "pConsultar"
End Sub

'*****************************************************************************************************
'* Nombre           : pImprimir
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPRIME LOS DATOS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pImprimir()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    X_PRINT.Imprimir_x_VSFlexGrid IIf(TabOne1.CurrTab = 0, Fg1, Fg2), T_RPT_TITULO + " ", "", T_RPT_PERIODO, True, True
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
    
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"
End Sub

Private Sub chk_Click(Index As Integer)
    chk(0).Enabled = True
    chk(1).Enabled = True
    If Index = 2 And chk(2).Value = 1 Then
        chk(0).Value = 0
        chk(1).Value = 0
        chk(0).Enabled = False
        chk(1).Enabled = False
        txtObs.Text = ""
    End If
End Sub

Private Sub CmdEditor_Click(Index As Integer)
    pHabilitarBotonEditor False
End Sub

Private Sub CmdOpcion_Click()
    pHabilitarBotonEditor True
End Sub

Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Row = 0 Then Exit Sub
   
    If NulosC(fg(Index).TextMatrix(Row, Col)) = "" Then fg(Index).TextMatrix(Row, 1) = ""
End Sub

Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    
    If validar_letras(KeyAscii) = False Then
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    On Error GoTo error
    Dim mTipoConsulta As Integer
    If SeEjecuto = True Then Exit Sub
    SeEjecuto = True
    
    cb_top.ListIndex = 0
    txt_top.Text = ""
    txtObs.Text = ""
    txtLote.Text = ""
    
    TxtFecha(0).valor = CDate("01/01/" + CStr(Year(Date)))
    TxtFecha(1).valor = Date
    ' detalle
    mTipoConsulta = fEstiloConsulta(1)
    fGenerarConsulta mTipoConsulta
    pConfigurarGrilla Fg1, mTipoConsulta
    ' resumen
    mTipoConsulta = fEstiloConsulta(2)
    fGenerarConsulta mTipoConsulta
    pConfigurarGrilla Fg2, mTipoConsulta
    Exit Sub

error:
    SHOW_ERROR Me.Name, "Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True ' interrumpir
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    On Error GoTo error
    Me.WindowState = 2
    Me.Height = 7950
    Me.Width = 11910
    
    SeEjecuto = False
    CentrarFrm Me
    LimpiaText txt_cb
    LimpiaText lbl_cb
    ' FORMATO DE LAS GRILLAS
    GRID_COMBOLIST fg(0), 2:        fg(0).Tag = fg(0).FormatString
    GRID_COMBOLIST fg(1), 2:        fg(1).Tag = fg(1).FormatString
    GRID_COMBOLIST fg(2), 2:        fg(2).Tag = fg(2).FormatString
    cb_top.AddItem "Primeros"
    cb_top.AddItem "Últimos"
    cb_top.ListIndex = 0
    Exit Sub

error:
    SHOW_ERROR
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    
    TabOne1.Width = Me.Width - 70
    TabOne1.Top = 2500
    If Me.Height > 2900 Then
        TabOne1.Height = Me.Height - 2900
    Else
        TabOne1.Height = Me.Height - 400
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BAND_INTERRUMPIR = True
    Erase ARR_TMP
End Sub

'*****************************************************************************************************
'* Nombre           : fValidarConsulta
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : FUNCION QUE VALIDARA LA CONSULTA DE LA FECHA ES NULL
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Private Function fValidarConsulta() As Boolean
    If TxtFecha(0).valor = "" Or TxtFecha(1).valor = "" Then
        MsgBox "Ingrese una fecha", vbExclamation, xTitulo
        If TxtFecha(0).valor = "" Then TxtFecha(0).SetFocus Else TxtFecha(1).SetFocus
        Exit Function
    End If
    
    If CDate(TxtFecha(0).valor) > CDate(TxtFecha(1).valor) Then
        MsgBox "La fecha inicial es superior al Final", vbExclamation, xTitulo
        TxtFecha(0).SetFocus
        Exit Function
    End If
    fValidarConsulta = True
End Function

'*****************************************************************************************************
'* Nombre           : fGenerarConsulta
'* Tipo             : FUNCION
'* Descripcion      : Generar la el Script SQL incluido filtros, establecer la cantidad de columnas,
'*                    a mostrar establecer el titulo del reporte, ESTA FUNCION DEVUELVE Script SQL
'*                    generado listo para su conexion, columnas definidas listo para generar la
'*                    estructura del grid
'* Paranetros       : NOMBRE         |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    mTipoConsulta  |  Integer    |  Tipo de Consulta, segun seleccion del Usuario
'* Devuelve         : String
'*****************************************************************************************************
Private Function fGenerarConsulta(mTipoConsulta As Integer) As String
    Dim nSQLSelect As String            ' CONSULTA GENERAL, ESTO PERMITIRA HACER LA CONSULTA
    Dim nSQL As String
    Dim nSQLFecha As String             ' almacenar el intervalo de fechas
    Dim nSQLProducto As String          ' almacenar los id's de productos
    Dim nSQLPersonal As String          ' almacenara los id's del personal
    Dim nSQLTarea As String             ' almacenara los id's de las tareas
    Dim nSQLArea As String              ' almacena el id del area
    Dim nSQLObs As String               ' almacena si se muestran solo los observados
    Dim nSQLInconsistenciaCab As String
    Dim nSQLInconsistenciaDet As String
    Dim k As Integer
    Dim nSQLOpcGral As String           ' almacenara los filtros de observados, reproceso, observaciones
    Dim nSQLOpcDet As String            ' almacenara los filtros de tipo de tarea
    
    ' de la fecha
    If CDate(TxtFecha(0).valor) < CDate(TxtFecha(1).valor) Then
        nSQLFecha = " ( pro_controltar.fchtra BETWEEN CDATE ('" + TxtFecha(0).valor + "') AND CDATE('" + TxtFecha(1).valor + "') ) "
        T_RPT_PERIODO = " Del: " + CStr(TxtFecha(0).valor) + " Al: " + CStr(TxtFecha(1).valor)
    Else
        nSQLFecha = " pro_controltar.fchtra = CDATE('" + TxtFecha(0).valor + "') "
         T_RPT_PERIODO = "Al: " + CStr(TxtFecha(1).valor)
   End If
    
    ' del personal
    nSQLPersonal = GENERAR_SQL_ID(fg(0), 1, " AND pla_empleados.id", "IN")
    
    ' de los productos
    nSQLProducto = GENERAR_SQL_ID(fg(1), 1, " AND alm_inventario.id", "IN")
    
    ' de las tareas
    nSQLTarea = GENERAR_SQL_ID(fg(2), 1, " AND pro_controltardet.idtar ", "IN")
    
    ' del area
    If NulosN(lbl_cod(0).Caption) <> 0 Then nSQLArea = " AND pro_controltar.idarea = " & NulosN(lbl_cod(0).Caption)
    
    ' solo los observados
    If chk(0).Value = 1 Then nSQLOpcGral = " AND pro_controltardet.observado = -1 "
    If chk(1).Value = 1 Then nSQLOpcGral = nSQLOpcGral & " AND pro_controltardet.reproceso = -1 "
    If NulosC(txtObs.Text) <> "" Then nSQLOpcGral = nSQLOpcGral & " AND pro_controltardet.observacion like '%" & NulosC(txtObs.Text) & "%' "
    If NulosC(txtLote.Text) <> "" Then nSQLOpcGral = nSQLOpcGral & " AND pro_controltardet.numlote like '%" & NulosC(txtLote.Text) & "%' "
    
    If OptTarea(0).Value = True Then
        nSQLOpcDet = ""
    ElseIf OptTarea(1).Value = True Then
        nSQLOpcDet = " AND pro_tareas.diverso = 0 "
    Else
        nSQLOpcDet = " AND pro_tareas.diverso = -1 "
    End If
    
    If chk(2).Value = 1 Then
        nSQLInconsistenciaCab = " AND ((pro_controltardet.idtar = 0 OR pro_controltardet.idtar is null ) OR ((pro_controltardet.idtar <> 0 AND pro_tareas.diverso =0  AND ( pro_controltardet.idrec is null or pro_controltardet.idrec = 0 ) AND (trim(pro_controltardet.observacion) ='' OR pro_controltardet.observacion is null) )) OR (pro_controltardet.horini is null OR pro_controltardet.horfin is null ) OR  (pro_controltardet.cant = 0 AND pro_controltardet.idunimed <> 7 ) ) "
        nSQLInconsistenciaDet = " AND ((pro_controltardet.idtar = 0 OR pro_controltardet.idtar is null ) OR ((pro_controltardet.idtar <> 0 AND pro_tareas.diverso =0  AND ( pro_controltardet.idrec is null or pro_controltardet.idrec = 0 ) AND (trim(pro_controltardet.observacion) ='' OR pro_controltardet.observacion is null) )) OR (pro_controltardetgr.horini is null OR pro_controltardetgr.horfin is null ) OR  (pro_controltardetgr.cant = 0 AND pro_controltardet.idunimed <> 7 ) ) "
    End If
    
    '--GENERAR LA CONSULTA SEGUN CONDICIONES
    Dim nSQLValor As String
    Dim nSQLCampos As String
    Dim nSQLWhere As String
    Dim nSQLFrom As String
    Dim nSQLGroupBy As String
    Dim nSQLOrderBy As String
    
    Q_COL_COMPARAR_FONDO = -1
    Select Case mTipoConsulta
        '************* CONTROL DE TAREAS
        ' detalle
        Case 0, 1, 2 ' x producto, x personal, x tarea
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 12:        Q_POSICION_TOTAL = 9:        Q_COL_COMPARAR_GRUPO = -1
            ' ADICIONAR DATOS AL GRID EN EL GRUPO (NOMBRE_GRUPO|COLUM1|COLUM2)
            Q_COL_GRUPO_ADD = -1:   N_CAMPO_GRUPO_ADD = ""
            Q_COL_GRUPO_INICIO = -1: Q_COL_GRUPO_TERMINA = -1
            Q_COL_COMPARAR_FONDO = 4
            T_RPT_TITULO = "DETALLE DEL REGISTRO DE TAREAS "
            If chk(0).Value = 1 Then T_RPT_TITULO = "DETALLE DEL REGISTRO DE TAREAS " & " - OBSERVADOS"
            
            nSQLCampos = "  [pro_controltardet].[idctr] & [pro_controltardet].[corr] AS codigo,pro_controltardet.numlote, pro_controltar.fchtra, mae_area.descripcion, IIf([pro_controltardet].[tipo]=1,[pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom],'Grupo Nº ' & [pro_controltardet].[idref]) AS nombres, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardet.horini, pro_controltardet.horfin, pro_controltardet.cant AS total, mae_unidades.abrev, IIf([pro_controltardet].[observado]=-1,'Si',' ') AS Obs, pro_controltardet.observacion "
            nSQLFrom = " ((pro_controltar INNER JOIN (alm_inventario RIGHT JOIN (((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr) LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id "
            nSQLWhere = nSQLFecha & nSQLArea & nSQLTarea & nSQLProducto & nSQLPersonal & nSQLOpcGral & nSQLOpcDet & nSQLInconsistenciaCab
            nSQLOrderBy = " pro_controltar.fchtra, IIf([pro_controltardet].[tipo]=1,[pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom],'Grupo Nº ' & [pro_controltardet].[idref]), pro_controltardet.horini;"
        
        ' resumen
        Case 9, 10, 11 ' x producto,x personal,x tarea
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 5:        Q_POSICION_TOTAL = 3:
            ' ADICIONAR DATOS AL GRID EN EL GRUPO (NOMBRE_GRUPO|COLUM1|COLUM2)
            Q_COL_GRUPO_ADD = -1:   N_CAMPO_GRUPO_ADD = ""
            Q_COL_GRUPO_INICIO = 1: Q_COL_GRUPO_TERMINA = 5
            If mTipoConsulta = 9 Then
                Q_COL_COMPARAR_GRUPO = 1
                Q_COL_COMPARAR_FONDO = 2 ' num lote
                T_RPT_TITULO = "RESUMEN DE TAREAS AGRUPADO POR PRODUCTO"
            ElseIf mTipoConsulta = 10 Then
                
            ElseIf mTipoConsulta = 11 Then
                Q_COL_COMPARAR_GRUPO = 3
                Q_COL_COMPARAR_FONDO = 2 ' num lote
                T_RPT_TITULO = "RESUMEN AGRUPADO POR TAREA"
            End If
            
            nSQLSelect = "SELECT vw.id , vw.producto,vw.numlote, vw.tarea, sum(vw.total) as acumulado ,vw.abrev " _
                + vbCr + " FROM ( "
            nSQLSelect = nSQLSelect _
                + vbCr + " SELECT alm_inventario.id ,IIf([alm_inventario].[descripcion] Is Not Null,[alm_inventario].[descripcion],'* ' & [pro_controltardet].[observacion]) AS producto, pro_controltardet.numlote, pro_tareas.descripcion AS tarea, Sum(pro_controltardet.cant) AS total, mae_unidades.abrev " _
                + vbCr + " FROM (pro_controltar INNER JOIN (alm_inventario RIGHT JOIN (((pro_controltardet INNER JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id " _
                + vbCr + " WHERE pro_controltardet.tipo=1 and pro_controltardet.cant <> 0 and " & nSQLFecha & nSQLArea & nSQLTarea & nSQLProducto & nSQLPersonal & nSQLOpcGral & nSQLOpcDet & nSQLInconsistenciaCab _
                + vbCr + " GROUP BY alm_inventario.id ,IIf([alm_inventario].[descripcion] Is Not Null,[alm_inventario].[descripcion],'* ' & [pro_controltardet].[observacion]), pro_controltardet.numlote, pro_tareas.descripcion, pro_tareas.descripcion, mae_unidades.abrev "
            nSQLSelect = nSQLSelect _
                + vbCr + " UNION "
            nSQLSelect = nSQLSelect _
                + vbCr + " SELECT alm_inventario.id ,IIf([alm_inventario].[descripcion] Is Not Null,[alm_inventario].[descripcion],'* ' & [pro_controltardet].[observacion]) AS producto, pro_controltardet.numlote, pro_tareas.descripcion AS tarea, Sum(pro_controltardetgr.cant) AS total, mae_unidades.abrev " _
                + vbCr + " FROM (pro_controltar INNER JOIN (alm_inventario RIGHT JOIN ((((pro_controltardet INNER JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) INNER JOIN pro_controltardetgr ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr) LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id " _
                + vbCr + " WHERE pro_controltardet.tipo=2 and  pro_controltardet.cant <> 0 and " & nSQLFecha & nSQLArea & nSQLTarea & nSQLProducto & nSQLPersonal & nSQLOpcGral & nSQLOpcDet & nSQLInconsistenciaDet _
                + vbCr + " GROUP BY alm_inventario.id ,IIf([alm_inventario].[descripcion] Is Not Null,[alm_inventario].[descripcion],'* ' & [pro_controltardet].[observacion]), pro_controltardet.numlote, pro_tareas.descripcion, mae_unidades.abrev "
            nSQLSelect = nSQLSelect _
                + vbCr + " ) AS vw " _
                + vbCr + " GROUP BY vw.id, producto, vw.numlote,vw.tarea,vw.abrev " _
                + vbCr + " ORDER BY vw.producto, vw.numlote asc,vw.tarea"
                        
        
        '************* EFICIENCIA
        ' detalle
        Case 3, 4, 5 ' x producto, x personal
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 17:        Q_POSICION_TOTAL = 6:
            
            If mTipoConsulta = 3 Then
                Q_COL_COMPARAR_GRUPO = 6 ' agrupado por producto
                T_RPT_TITULO = "EFICIENCIA DEL PERSONAL AGRUPADO POR PRODUCTO"
                Q_COL_COMPARAR_FONDO = 1
            ElseIf mTipoConsulta = 4 Then
                Q_COL_COMPARAR_GRUPO = 4 ' agrupado por personal
                T_RPT_TITULO = "EFICIENCIA DEL PERSONAL AGRUPADO POR PERSONAL"
                Q_COL_GRUPO_INICIO = 6: Q_COL_GRUPO_TERMINA = 10
                Q_COL_COMPARAR_FONDO = 2
            ElseIf mTipoConsulta = 5 Then
                Q_COL_COMPARAR_GRUPO = 5 ' agrupado por tarea
                T_RPT_TITULO = "EFICIENCIA DEL PERSONAL AGRUPADO POR TAREA"
                Q_COL_GRUPO_INICIO = 6: Q_COL_GRUPO_TERMINA = 10
                Q_COL_COMPARAR_FONDO = 2
            End If
            
            Q_COL_GRUPO_INICIO = 2: Q_COL_GRUPO_TERMINA = 10
            ' ADICIONAR DATOS AL GRID EN EL GRUPO (NOMBRE_GRUPO|COLUM1|COLUM2)
            Q_COL_GRUPO_ADD = -1:   N_CAMPO_GRUPO_ADD = ""
            
            nSQLSelect = "SELECT vwtarea.*, vwcosto.canteo, iif(vwtarea.unidxhor = 0 or vwcosto.canteo = 0 or vwcosto.canteo is null ,0, (vwtarea.unidxhor/vwcosto.canteo)*100 ) as Eficiencia " _
                + vbCr + " FROM ( "
            nSQLSelect = nSQLSelect _
                + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk, pro_controltardet.numlote, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardet.observacion, pro_controltardet.horini, pro_controltardet.horfin, " _
                + vbCr + " IIf(pro_controltardet.cant Is Null,0,CDbl(pro_controltardet.cant)) AS CantReal, mae_unidades.abrev, IIf([pro_controltardet].[horini] Is Null Or [pro_controltardet].[horfin] Is Null,'',IIF([pro_controltardet].[horini]<cdate('13:20:00') AND [pro_controltardet].[horfin]>cdate('14:00:00'),format(cdate(format([pro_controltardet].[horfin]-[pro_controltardet].[horini],'hh:mm:ss')) - cdate('01:00:00'),'hh:mm:ss'),Format([pro_controltardet].[horfin]-[pro_controltardet].[horini],'hh:mm:ss'))) AS difhora, IIf([difhora] Is Null Or [difhora]='',0,Hour([difhora])*60+Minute([difhora])) AS totmin, IIf([CantReal]=0 Or [totmin]=0,0,[CantReal]/[totmin]) AS UnidXMin, [UnidXMin]*60 AS UnidXHor " _
                + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (alm_inventario RIGHT JOIN ((((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr " _
                + vbCr + " WHERE pro_controltardet.tipo=1 AND  pro_controltardet.cant<>0 AND (pro_controltardet.idtar <>0 OR pro_controltardet.idrec <>0) AND " & nSQLFecha & nSQLArea & nSQLPersonal & nSQLTarea & nSQLProducto & nSQLPersonal & nSQLOpcGral & nSQLOpcDet & nSQLInconsistenciaCab
            nSQLSelect = nSQLSelect _
                + vbCr + " UNION "
            nSQLSelect = nSQLSelect _
                + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk, pro_controltardet.numlote, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardet.observacion, pro_controltardetgr.horini, pro_controltardetgr.horfin, " _
                + vbCr + " IIf(pro_controltardetgr.cant Is Null,0,CDbl(pro_controltardetgr.cant)) AS CantReal, mae_unidades.abrev, IIf([pro_controltardetgr].[horini] Is Null Or [pro_controltardetgr].[horfin] Is Null,'',IIF([pro_controltardetgr].[horini]<cdate('13:20:00') AND [pro_controltardetgr].[horfin]>cdate('14:00:00'),format(cdate(format([pro_controltardetgr].[horfin]-[pro_controltardetgr].[horini],'hh:mm:ss')) - cdate('01:00:00'),'hh:mm:ss'),Format([pro_controltardetgr].[horfin]-[pro_controltardetgr].[horini],'hh:mm:ss'))) AS difhora, IIf([difhora] Is Null Or [difhora]='',0,Hour([difhora])*60+Minute([difhora])) AS totmin, IIf([CantReal]=0 Or [totmin]=0,0,[CantReal]/[totmin]) AS UnidXMin, [UnidXMin]*60 AS UnidXHor " _
                + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (alm_inventario RIGHT JOIN ((((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) INNER JOIN (pro_controltardetgr LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id) ON (pro_controltardet.idctr = pro_controltardetgr.idctr) AND (pro_controltardet.corr = pro_controltardetgr.corr)) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr " _
                + vbCr + " WHERE pro_controltardet.tipo=2 AND  pro_controltardetgr.cant<>0 AND (pro_controltardet.idtar <>0 OR pro_controltardet.idrec <>0) AND pro_controltardetgr.activo = -1 AND " & nSQLFecha & nSQLArea & nSQLPersonal & nSQLTarea & nSQLProducto & nSQLPersonal & nSQLOpcGral & nSQLOpcDet & nSQLInconsistenciaDet
            nSQLSelect = nSQLSelect _
                + vbCr + " ) AS vwtarea "
            nSQLSelect = nSQLSelect _
                + vbCr + " Left Join " _
                + vbCr + " (SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden " _
                + vbCr + " FROM pro_tareas INNER JOIN ((alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
                + vbCr + " ) AS vwcosto"
            nSQLSelect = nSQLSelect _
                + vbCr + " ON vwtarea.codigopk = vwcosto.codigopk "
            
            If mTipoConsulta = 3 Then ' agrup. producto
                nSQLSelect = nSQLSelect _
                    + vbCr + " ORDER BY vwtarea.producto,vwtarea.numlote,vwtarea.fchtra, vwtarea.area,vwtarea.personal,vwtarea.horini "
            ElseIf mTipoConsulta = 4 Then ' agrup. personal
                nSQLSelect = nSQLSelect _
                    + vbCr + " ORDER BY vwtarea.personal,vwtarea.fchtra,vwtarea.area,vwtarea.horini "
            Else ' agrup. tarea
                nSQLSelect = nSQLSelect _
                    + vbCr + " ORDER BY vwtarea.tarea,vwtarea.fchtra,vwtarea.area,vwtarea.horini "
            End If
            
        ' resumen - eficiencia
        Case 12 ' x producto

        Case 13 ' x personal
        
        Case 14 ' x tarea
        
        '************* MAXIMOS Y MINIMOS
        ' detalle
        Case 6, 7, 8 ' x producto, x personal,x tarea
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 17:        Q_POSICION_TOTAL = 6:
            
            If mTipoConsulta = 6 Then       ' x producto
                Q_COL_COMPARAR_GRUPO = 6    ' agrupado por producto
                T_RPT_TITULO = "DETALLE DE MAXIMOS Y MINIMOS AGRUPADOS POR PRODUCTO"
                Q_COL_COMPARAR_FONDO = 2
            ElseIf mTipoConsulta = 7 Then   ' x personal
                Q_COL_COMPARAR_GRUPO = 4    ' agrupado por producto
                T_RPT_TITULO = "DETALLE DE MAXIMOS Y MINIMOS AGRUPADOS POR PERSONAL"
                Q_COL_COMPARAR_FONDO = 2
            ElseIf mTipoConsulta = 8 Then   ' x tarea
                Q_COL_COMPARAR_GRUPO = 5    ' agrupado tarea
                T_RPT_TITULO = "DETALLE DE MAXIMOS Y MINIMOS AGRUPADOS POR TAREA"
                Q_COL_GRUPO_INICIO = 6: Q_COL_GRUPO_TERMINA = 10
                Q_COL_COMPARAR_FONDO = 2
            End If
            
            Q_COL_GRUPO_INICIO = 2: Q_COL_GRUPO_TERMINA = 7
            
            ' ADICIONAR DATOS AL GRID EN EL GRUPO (NOMBRE_GRUPO|COLUM1|COLUM2)
            Q_COL_GRUPO_ADD = -1:   N_CAMPO_GRUPO_ADD = ""
            nSQLSelect = "SELECT vwtarea.*, vwcosto.canteo, iif(vwtarea.unidxhor = 0 or vwcosto.canteo = 0 or vwcosto.canteo is null ,0, (vwtarea.unidxhor/vwcosto.canteo)*100 ) as Eficiencia " _
                + vbCr + " FROM ( "
            nSQLSelect = nSQLSelect _
                + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk, pro_controltardet.numlote, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardet.observacion, pro_controltardet.horini, pro_controltardet.horfin, " _
                + vbCr + " IIf(pro_controltardet.cant Is Null,0,CDbl(pro_controltardet.cant)) AS CantReal, mae_unidades.abrev, IIf([pro_controltardet].[horini] Is Null Or [pro_controltardet].[horfin] Is Null,'',IIF([pro_controltardet].[horini]<cdate('13:20:00') AND [pro_controltardet].[horfin]>cdate('14:00:00'),format(cdate(format([pro_controltardet].[horfin]-[pro_controltardet].[horini],'hh:mm:ss')) - cdate('01:00:00'),'hh:mm:ss'),Format([pro_controltardet].[horfin]-[pro_controltardet].[horini],'hh:mm:ss'))) AS difhora, IIf([difhora] Is Null Or [difhora]='',0,Hour([difhora])*60+Minute([difhora])) AS totmin, IIf([CantReal]=0 Or [totmin]=0,0,[CantReal]/[totmin]) AS UnidXMin, [UnidXMin]*60 AS UnidXHor " _
                + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (alm_inventario RIGHT JOIN ((((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr " _
                + vbCr + " WHERE pro_controltardet.tipo=1 AND  pro_controltardet.cant<>0 AND (pro_controltardet.idtar <>0 OR pro_controltardet.idrec <>0) AND " & nSQLFecha & nSQLArea & nSQLPersonal & nSQLTarea & nSQLProducto & nSQLPersonal & nSQLOpcGral & nSQLOpcDet & nSQLInconsistenciaCab
            nSQLSelect = nSQLSelect _
                + vbCr + " UNION "
            nSQLSelect = nSQLSelect _
                + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk, pro_controltardet.numlote, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardet.observacion, pro_controltardetgr.horini, pro_controltardetgr.horfin, " _
                + vbCr + " IIf(pro_controltardetgr.cant Is Null,0,CDbl(pro_controltardetgr.cant)) AS CantReal, mae_unidades.abrev, IIf([pro_controltardetgr].[horini] Is Null Or [pro_controltardetgr].[horfin] Is Null,'',IIF([pro_controltardetgr].[horini]<cdate('13:20:00') AND [pro_controltardetgr].[horfin]>cdate('14:00:00'),format(cdate(format([pro_controltardetgr].[horfin]-[pro_controltardetgr].[horini],'hh:mm:ss')) - cdate('01:00:00'),'hh:mm:ss'),Format([pro_controltardetgr].[horfin]-[pro_controltardetgr].[horini],'hh:mm:ss'))) AS difhora, IIf([difhora] Is Null Or [difhora]='',0,Hour([difhora])*60+Minute([difhora])) AS totmin, IIf([CantReal]=0 Or [totmin]=0,0,[CantReal]/[totmin]) AS UnidXMin, [UnidXMin]*60 AS UnidXHor " _
                + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (alm_inventario RIGHT JOIN ((((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) INNER JOIN (pro_controltardetgr LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id) ON (pro_controltardet.idctr = pro_controltardetgr.idctr) AND (pro_controltardet.corr = pro_controltardetgr.corr)) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr " _
                + vbCr + " WHERE pro_controltardet.tipo=2 AND  pro_controltardetgr.cant<>0 AND (pro_controltardet.idtar <>0 OR pro_controltardet.idrec <>0) AND pro_controltardetgr.activo = -1 AND " & nSQLFecha & nSQLArea & nSQLPersonal & nSQLTarea & nSQLProducto & nSQLPersonal & nSQLOpcGral & nSQLOpcDet & nSQLInconsistenciaDet
            nSQLSelect = nSQLSelect _
                + vbCr + " ) AS vwtarea "
            nSQLSelect = nSQLSelect _
                + vbCr + " Left Join " _
                + vbCr + " (SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden " _
                + vbCr + " FROM pro_tareas INNER JOIN ((alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
                + vbCr + " ) AS vwcosto"
            nSQLSelect = nSQLSelect _
                + vbCr + " ON vwtarea.codigopk = vwcosto.codigopk "
            
            If mTipoConsulta = 6 Then       ' agrup. producto
                nSQLSelect = nSQLSelect _
                    + vbCr + " ORDER BY vwtarea.producto,vwtarea.fchtra, vwtarea.personal,vwtarea.area,vwtarea.horini"
            ElseIf mTipoConsulta = 7 Then   ' agrup. personal
                nSQLSelect = nSQLSelect _
                    + vbCr + " ORDER BY vwtarea.personal, vwtarea.fchtra, vwtarea.area,vwtarea.horini"
            ElseIf mTipoConsulta = 8 Then   ' agrup. tarea
                nSQLSelect = nSQLSelect _
                    + vbCr + " ORDER BY vwtarea.tarea,vwtarea.fchtra,vwtarea.personal,  vwtarea.area,vwtarea.horini"
            End If
                        
        ' resumen - maximos y minimos
        Case 15, 16, 17 ' x producto,x personal, x tarea
            Q_COL_FILA_OCULTA = -1:         Q_COL_FILA = 13:        Q_POSICION_TOTAL = 3:
            Q_COL_GRUPO_INICIO = 1: Q_COL_GRUPO_TERMINA = 2
            
            If mTipoConsulta = 15 Then
                Q_COL_COMPARAR_GRUPO = 1    ' agrupado por producto
                T_RPT_TITULO = "RESUMEN DE MAXIMOS Y MINIMOS AGRUPADO POR PRODUCTO"
                Q_COL_COMPARAR_FONDO = 1
            ElseIf mTipoConsulta = 16 Then
                Q_COL_COMPARAR_GRUPO = 2    ' agrupado por personal
                T_RPT_TITULO = "RESUMEN DE MAXIMOS Y MINIMOS AGRUPADO POR PERSONAL"
                Q_COL_COMPARAR_FONDO = 1
            ElseIf mTipoConsulta = 17 Then
                Q_COL_COMPARAR_GRUPO = 0    ' agrupado tarea
                T_RPT_TITULO = "RESUMEN DE MAXIMOS Y MINIMOS AGRUPADO POR TAREA"
                Q_COL_COMPARAR_FONDO = 2
            End If

            ' ADICIONAR DATOS AL GRID EN EL GRUPO (NOMBRE_GRUPO|COLUM1|COLUM2)
            Q_COL_GRUPO_ADD = -1:   N_CAMPO_GRUPO_ADD = ""
            nSQLSelect = "SELECT vw.tarea, vw.producto," & IIf(mTipoConsulta = 16, "vw.personal", "''") & " AS personal, Sum(vw.CantReal) AS total, vw.abrev, Min(vw.UnidXHor) AS CanHrMin, Avg(vw.UnidXHor) AS CanHrProm, Max(vw.UnidXHor) AS CanHrMax, Min(vw.Eficiencia) AS EficMin, Avg(vw.Eficiencia) AS EficProm, Max(vw.Eficiencia) AS EficMax, Count(vw.codigopk) AS CanReg " _
                + vbCr + " FROM ( "

            nSQLSelect = nSQLSelect _
                + vbCr + " SELECT vwtarea.*, vwcosto.canteo, iif(vwtarea.unidxhor = 0 or vwcosto.canteo = 0 or vwcosto.canteo is null ,0, (vwtarea.unidxhor/vwcosto.canteo)*100 ) as Eficiencia " _
                    + vbCr + " FROM ( "
                    nSQLSelect = nSQLSelect _
                        + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk, pro_controltardet.numlote, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardet.observacion, pro_controltardet.horini, pro_controltardet.horfin, " _
                        + vbCr + " IIf(pro_controltardet.cant Is Null,0,CDbl(pro_controltardet.cant)) AS CantReal, mae_unidades.abrev, IIf([pro_controltardet].[horini] Is Null Or [pro_controltardet].[horfin] Is Null,'',IIF([pro_controltardet].[horini]<cdate('13:20:00') AND [pro_controltardet].[horfin]>cdate('14:00:00'),format(cdate(format([pro_controltardet].[horfin]-[pro_controltardet].[horini],'hh:mm:ss')) - cdate('01:00:00'),'hh:mm:ss'),Format([pro_controltardet].[horfin]-[pro_controltardet].[horini],'hh:mm:ss'))) AS difhora, IIf([difhora] Is Null Or [difhora]='',0,Hour([difhora])*60+Minute([difhora])) AS totmin, IIf([CantReal]=0 Or [totmin]=0,0,[CantReal]/[totmin]) AS UnidXMin, [UnidXMin]*60 AS UnidXHor " _
                        + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (alm_inventario RIGHT JOIN ((((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr " _
                        + vbCr + " WHERE pro_controltardet.tipo=1 AND pro_controltardet.cant<>0 AND (pro_controltardet.idtar <>0 OR pro_controltardet.idrec <>0) AND " & nSQLFecha & nSQLArea & nSQLPersonal & nSQLTarea & nSQLProducto & nSQLPersonal & nSQLOpcGral & nSQLOpcDet & nSQLInconsistenciaCab
                    nSQLSelect = nSQLSelect _
                        + vbCr + " UNION "
                    nSQLSelect = nSQLSelect _
                        + vbCr + " SELECT IIf([pro_controltardet].[idrec]=0 Or [pro_controltardet].[idrec] Is Null,'-',[pro_controltardet].[idrec]) & '*' & [pro_controltardet].[idtar] & '*' & [pro_controltardet].[idunimed] AS codigopk, pro_controltardet.numlote, pro_controltar.fchtra, mae_area.descripcion AS area, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS personal, pro_tareas.descripcion AS tarea, alm_inventario.descripcion AS producto, pro_controltardet.observacion, pro_controltardetgr.horini, pro_controltardetgr.horfin, " _
                        + vbCr + " IIf(pro_controltardetgr.cant Is Null,0,CDbl(pro_controltardetgr.cant)) AS CantReal, mae_unidades.abrev, IIf([pro_controltardetgr].[horini] Is Null Or [pro_controltardetgr].[horfin] Is Null,'',IIF([pro_controltardetgr].[horini]<cdate('13:20:00') AND [pro_controltardetgr].[horfin]>cdate('14:00:00'),format(cdate(format([pro_controltardetgr].[horfin]-[pro_controltardetgr].[horini],'hh:mm:ss')) - cdate('01:00:00'),'hh:mm:ss'),Format([pro_controltardetgr].[horfin]-[pro_controltardetgr].[horini],'hh:mm:ss'))) AS difhora, IIf([difhora] Is Null Or [difhora]='',0,Hour([difhora])*60+Minute([difhora])) AS totmin, IIf([CantReal]=0 Or [totmin]=0,0,[CantReal]/[totmin]) AS UnidXMin, [UnidXMin]*60 AS UnidXHor " _
                        + vbCr + " FROM (pro_controltar LEFT JOIN mae_area ON pro_controltar.idarea = mae_area.id) INNER JOIN (alm_inventario RIGHT JOIN ((((pro_controltardet LEFT JOIN pro_tareas ON pro_controltardet.idtar = pro_tareas.id) LEFT JOIN pro_receta ON pro_controltardet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_controltardet.idunimed = mae_unidades.id) INNER JOIN (pro_controltardetgr LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id) ON (pro_controltardet.idctr = pro_controltardetgr.idctr) AND (pro_controltardet.corr = pro_controltardetgr.corr)) ON alm_inventario.id = pro_receta.iditem) ON pro_controltar.id = pro_controltardet.idctr " _
                        + vbCr + " WHERE pro_controltardet.tipo=2 AND pro_controltardetgr.cant<>0 AND (pro_controltardet.idtar <>0 OR pro_controltardet.idrec <>0) AND pro_controltardetgr.activo = -1 AND " & nSQLFecha & nSQLArea & nSQLPersonal & nSQLTarea & nSQLProducto & nSQLPersonal & nSQLOpcGral & nSQLOpcDet & nSQLInconsistenciaDet
                    nSQLSelect = nSQLSelect _
                        + vbCr + " ) AS vwtarea "
                nSQLSelect = nSQLSelect _
                    + vbCr + " Left Join " _
                        + vbCr + " (SELECT IIf([pro_costo].[tipo]=1,[pro_costo].[idref],'-') & '*' & [pro_costodet].[idtar] & '*' & [pro_costodet].[idunimed] AS codigopk, IIf([pro_costo].[tipo]=1,'Receta','Tarea Diversa') AS origen, IIf([pro_costo].[tipo]=1,[alm_inventario].[descripcion],[pro_tareas_1].[descripcion]) AS referencia, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.minimo, pro_costodet.maximo, pro_costodet.promedio, pro_costodet.jornal, pro_costodet.cant AS canteo, pro_costodet.costo, pro_costodet.orden " _
                        + vbCr + " FROM pro_tareas INNER JOIN ((alm_inventario RIGHT JOIN ((pro_costo LEFT JOIN pro_receta ON pro_costo.idref = pro_receta.id) LEFT JOIN pro_tareas AS pro_tareas_1 ON pro_costo.idref = pro_tareas_1.id) ON alm_inventario.id = pro_receta.iditem) INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
                        + vbCr + " ) AS vwcosto"
                nSQLSelect = nSQLSelect _
                    + vbCr + " ON vwtarea.codigopk = vwcosto.codigopk "
                
            nSQLSelect = nSQLSelect _
                + vbCr + " ) AS vw "
                
            If mTipoConsulta = 15 Then     ' agrup. producto
                nSQLSelect = nSQLSelect _
                    + vbCr + " GROUP BY  vw.producto, vw.tarea, vw.abrev; "
                                        
            ElseIf mTipoConsulta = 16 Then ' agrup. personal
                nSQLSelect = nSQLSelect _
                    + vbCr + " GROUP BY vw.personal, vw.producto, vw.tarea, vw.abrev; "
                
            ElseIf mTipoConsulta = 17 Then ' agrup. tarea
                nSQLSelect = nSQLSelect _
                    + vbCr + " GROUP BY vw.tarea, vw.producto, vw.abrev; "
            End If
    End Select
    
    ' DE LA CANTIDAD DE COL
    Q_COL_FILA = Q_COL_FILA + 1
    Q_POS_MES_INICIO = Q_COL_FILA          ' Q_COL_FILA + CAMPO_TOTAL

    ' GENERANDO LA CONSULTA
    If nSQLSelect = "" Then
        nSQLSelect = "SELECT " + nSQLCampos + _
        vbCr + " FROM " + nSQLFrom + _
        vbCr + " WHERE " + nSQLWhere + _
        vbCr + IIf(nSQLGroupBy <> "", " GROUP BY ", "") + nSQLGroupBy + _
        vbCr + " ORDER BY " + nSQLOrderBy
    End If
    fGenerarConsulta = nSQLSelect
End Function

Private Sub Limpiar_ARRAY_TOTAL(Optional F_LIMPIA_TOT_GRL As Boolean = False)
    Dim k As Integer
    For k = 0 To UBound(ARR_TMP())
'        ARR_TMP(k, 3) = 0
'        If F_LIMPIA_TOT_GRL = True Then ARR_TMP(k, 4) = 0
    Next
End Sub

'*****************************************************************************************************
'* Nombre           : pConfigurarGrilla
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : Establecer los encabezados del grid, DEVUELVE LA Grilla con Encabezado
'* Paranetros       : NOMBRE         |  TIPO          |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Grid           |  VSFlexGrid    |  Es el objeto VSFlexGrid
'*                    mTipoConsulta  |  Integer       |  valor segun la consulta seleccionada
'* Devuelve         :
'*****************************************************************************************************
Private Sub pConfigurarGrilla(Grid As VSFlexGrid, mTipoConsulta As Integer)
    Dim k As Integer
    
    Grid.FrozenCols = 0

    With Grid
        .Cols = 1
        .Cols = Q_COL_FILA_OCULTA + Q_COL_FILA
        Q_POS_MES = Q_POS_MES_INICIO
        .ColWidth(0) = 200
        
        Select Case mTipoConsulta
            '************* CONTROL DE TAREAS
            ' detalle
            Case 0, 1, 2 ' x producto, x personal, x tarea
                    .TextMatrix(0, 2) = "Nº Lote":          .ColWidth(2) = 1200:      .ColAlignment(2) = flexAlignLeftBottom:         .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 3) = "Fecha":            .ColWidth(3) = 800:       .ColAlignment(3) = flexAlignCenterBottom:       .Row = 0: .Col = 3: .CellAlignment = flexAlignCenterBottom
                    .TextMatrix(0, 4) = "Area":             .ColWidth(4) = 700:       .ColAlignment(4) = flexAlignLeftBottom:         .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 5) = "Personal / Grupo": .ColWidth(5) = 1380:      .ColAlignment(5) = flexAlignLeftBottom:         .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 6) = "Tarea":            .ColWidth(6) = 2800:      .ColAlignment(6) = flexAlignLeftBottom:         .Row = 0: .Col = 6: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 7) = "Producto":         .ColWidth(7) = 3000:      .ColAlignment(7) = flexAlignLeftBottom:         .Row = 0: .Col = 7: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 8) = "H.Inicio":         .ColWidth(8) = 800:       .ColAlignment(8) = flexAlignCenterCenter:       .Row = 0: .Col = 8: .CellAlignment = flexAlignCenterCenter
                    .TextMatrix(0, 9) = "H.Final":          .ColWidth(9) = 800:       .ColAlignment(9) = flexAlignCenterCenter:       .Row = 0: .Col = 9: .CellAlignment = flexAlignCenterCenter
                    .TextMatrix(0, 10) = "Cant":            .ColWidth(10) = 850:      .ColAlignment(10) = flexAlignRightBottom:       .Row = 0: .Col = 10: .CellAlignment = flexAlignRightBottom
                    .TextMatrix(0, 11) = "U.M.":            .ColWidth(11) = 450:      .ColAlignment(11) = flexAlignCenterCenter:      .Row = 0: .Col = 11: .CellAlignment = flexAlignCenterCenter
                    .TextMatrix(0, 12) = "Obs":             .ColWidth(12) = 400:      .ColAlignment(12) = flexAlignLeftBottom:        .Row = 0: .Col = 12: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 13) = "Inf. Adicional":  .ColWidth(13) = 2000:     .ColAlignment(13) = flexAlignLeftBottom:        .Row = 0: .Col = 13: .CellAlignment = flexAlignLeftBottom
                    ' ocultar la columna de obs si solo se muestran los observados
                    If chk(0).Value = 1 Then
                        .ColWidth(12) = 400
                        .ColWidth(13) = 4000
                    End If
            
            ' resumen
            Case 9, 10, 11 ' x producto, x personal,x num lote
                    .TextMatrix(0, 2) = "Producto":     .ColWidth(2) = 4500:  .ColAlignment(2) = flexAlignLeftBottom:   .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 3) = "Nº Lote":      .ColWidth(3) = 1500:  .ColAlignment(3) = flexAlignLeftBottom:   .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 4) = "Tarea":        .ColWidth(4) = 4500:  .ColAlignment(4) = flexAlignLeftBottom:      .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 5) = "Total":        .ColWidth(5) = 1300:  .ColAlignment(5) = flexAlignRightBottom:  .Row = 0: .Col = 5: .CellAlignment = flexAlignRightCenter
                    .TextMatrix(0, 6) = "U.M.":         .ColWidth(6) = 450:   .ColAlignment(6) = flexAlignCenterCenter: .Row = 0: .Col = 6: .CellAlignment = flexAlignCenterCenter
                    
                    If mTipoConsulta = 9 Then .ColWidth(2) = 0  ' agrupado por producto
                    If mTipoConsulta = 11 Then .ColWidth(4) = 0 ' agrupado por personal

            '************* EFICIENCIA
            ' detalle
            Case 3, 4, 5 ' x producto, x personal
                    .TextMatrix(0, 2) = "Nº Lote":          .ColWidth(2) = 1200:      .ColAlignment(2) = flexAlignLeftBottom:         .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 3) = "Fecha":            .ColWidth(3) = 800:       .ColAlignment(3) = flexAlignCenterBottom:       .Row = 0: .Col = 3: .CellAlignment = flexAlignCenterBottom
                    .TextMatrix(0, 4) = "Area":             .ColWidth(4) = 700:       .ColAlignment(4) = flexAlignLeftBottom:         .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 5) = "Personal":         .ColWidth(5) = 1380:      .ColAlignment(5) = flexAlignLeftBottom:         .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 6) = "Tarea":            .ColWidth(6) = 2800:     .ColAlignment(6) = flexAlignLeftBottom:          .Row = 0: .Col = 6: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 7) = "Producto":         .ColWidth(7) = 3000:     .ColAlignment(7) = flexAlignLeftBottom:          .Row = 0: .Col = 7: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 8) = "Inf. Adicional":   .ColWidth(8) = 1500:     .ColAlignment(8) = flexAlignLeftBottom:          .Row = 0: .Col = 8: .CellAlignment = flexAlignLeftBottom
                    
                    .TextMatrix(0, 9) = "H.Inicio":         .ColWidth(9) = 800:      .ColAlignment(9) = flexAlignCenterCenter:         .Row = 0: .Col = 9: .CellAlignment = flexAlignCenterCenter
                    .TextMatrix(0, 10) = "H.Final":         .ColWidth(10) = 800:     .ColAlignment(10) = flexAlignCenterCenter:        .Row = 0: .Col = 10: .CellAlignment = flexAlignCenterCenter
                    .TextMatrix(0, 11) = "CantReal":        .ColWidth(11) = 800:     .ColAlignment(11) = flexAlignRightCenter:         .Row = 0: .Col = 11: .CellAlignment = flexAlignRightCenter
                    .TextMatrix(0, 12) = "U.M.":            .ColWidth(12) = 450:     .ColAlignment(12) = flexAlignCenterCenter:        .Row = 0: .Col = 12: .CellAlignment = flexAlignCenterCenter
                    
                    .TextMatrix(0, 13) = "Dif.Hora":        .ColWidth(13) = 750:     .ColAlignment(13) = flexAlignRightCenter:        .Row = 0: .Col = 13: .CellAlignment = flexAlignRightCenter
                    .TextMatrix(0, 14) = "Tot.Min":         .ColWidth(14) = 700:     .ColAlignment(14) = flexAlignRightCenter:        .Row = 0: .Col = 14: .CellAlignment = flexAlignRightCenter
                    .TextMatrix(0, 15) = "Unid x Min":      .ColWidth(15) = 0:       .ColAlignment(15) = flexAlignRightCenter:        .Row = 0: .Col = 15: .CellAlignment = flexAlignRightCenter
                    .TextMatrix(0, 16) = "Unid x Hor":      .ColWidth(16) = 1100:    .ColAlignment(16) = flexAlignRightCenter:        .Row = 0: .Col = 16: .CellAlignment = flexAlignRightCenter
                    .TextMatrix(0, 17) = "Cant Teo":        .ColWidth(17) = 850:     .ColAlignment(17) = flexAlignRightCenter:        .Row = 0: .Col = 17: .CellAlignment = flexAlignRightCenter
                    .TextMatrix(0, 18) = "Eficiencia":      .ColWidth(18) = 800:     .ColAlignment(18) = flexAlignRightCenter:        .Row = 0: .Col = 18: .CellAlignment = flexAlignRightCenter
                    
                    If mTipoConsulta = 3 Then .ColWidth(7) = 0 ' agrupado por producto
                    If mTipoConsulta = 4 Then .ColWidth(5) = 0 ' agrupado por personal
                    If mTipoConsulta = 5 Then .ColWidth(6) = 0 ' agrupado por tarea
                    
            ' resumen - eficiencia
            Case 12, 13, 14 ' x producto, x personal,x tarea
            
            '************* MAXIMOS Y MINIMOS
            ' detalle
            Case 6, 7, 8 ' x producto,x personal,x tarea
                    .FrozenCols = 7
                    .TextMatrix(0, 2) = "Nº Lote":          .ColWidth(2) = 0:        .ColAlignment(2) = flexAlignLeftBottom:          .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 3) = "Fecha":            .ColWidth(3) = 800:      .ColAlignment(3) = flexAlignCenterBottom:        .Row = 0: .Col = 3: .CellAlignment = flexAlignCenterBottom
                    .TextMatrix(0, 4) = "Area":             .ColWidth(4) = 700:      .ColAlignment(4) = flexAlignLeftBottom:          .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 5) = "Personal":         .ColWidth(5) = 1380:     .ColAlignment(5) = flexAlignLeftBottom:          .Row = 0: .Col = 5: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 6) = "Tarea":            .ColWidth(6) = 2800:     .ColAlignment(6) = flexAlignLeftBottom:          .Row = 0: .Col = 6: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 7) = "Producto":         .ColWidth(7) = 3000:     .ColAlignment(7) = flexAlignLeftBottom:          .Row = 0: .Col = 7: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 8) = "Inf. Adicional":   .ColWidth(8) = 900:      .ColAlignment(8) = flexAlignLeftBottom:          .Row = 0: .Col = 8: .CellAlignment = flexAlignLeftBottom
                    .TextMatrix(0, 9) = "H.Inicio":         .ColWidth(9) = 800:      .ColAlignment(9) = flexAlignCenterCenter:        .Row = 0: .Col = 9: .CellAlignment = flexAlignCenterCenter
                    .TextMatrix(0, 10) = "H.Final":         .ColWidth(10) = 800:     .ColAlignment(10) = flexAlignCenterCenter:       .Row = 0: .Col = 10: .CellAlignment = flexAlignCenterCenter
                    .TextMatrix(0, 11) = "CantReal":        .ColWidth(11) = 800:     .ColAlignment(11) = flexAlignRightCenter:        .Row = 0: .Col = 11: .CellAlignment = flexAlignRightCenter
                    .TextMatrix(0, 12) = "U.M.":            .ColWidth(12) = 450:     .ColAlignment(12) = flexAlignCenterCenter:       .Row = 0: .Col = 12: .CellAlignment = flexAlignCenterCenter
                    .TextMatrix(0, 13) = "Dif.Hora":        .ColWidth(13) = 0:       .ColAlignment(13) = flexAlignRightCenter:        .Row = 0: .Col = 13: .CellAlignment = flexAlignRightCenter
                    .TextMatrix(0, 14) = "Tot.Min":         .ColWidth(14) = 0:       .ColAlignment(14) = flexAlignRightCenter:        .Row = 0: .Col = 14: .CellAlignment = flexAlignRightCenter
                    .TextMatrix(0, 15) = "Unid x Min":      .ColWidth(15) = 0:       .ColAlignment(15) = flexAlignRightCenter:        .Row = 0: .Col = 15: .CellAlignment = flexAlignRightCenter
                    .TextMatrix(0, 16) = "Unid x Hor":      .ColWidth(16) = 1000:    .ColAlignment(16) = flexAlignRightCenter:        .Row = 0: .Col = 16: .CellAlignment = flexAlignRightCenter
                    .TextMatrix(0, 17) = "Cant Teo":        .ColWidth(17) = 850:     .ColAlignment(17) = flexAlignRightCenter:        .Row = 0: .Col = 17: .CellAlignment = flexAlignRightCenter
                    .TextMatrix(0, 18) = "Eficiencia":      .ColWidth(18) = 800:     .ColAlignment(18) = flexAlignRightCenter:        .Row = 0: .Col = 18: .CellAlignment = flexAlignRightCenter
                    
                    If mTipoConsulta = 6 Then .ColWidth(7) = 0 ' agrupado por producto
                    If mTipoConsulta = 7 Then .ColWidth(5) = 0 ' agrupado por personal
                    If mTipoConsulta = 8 Then .ColWidth(6) = 0 ' agrupado por tarea
                                        
            ' resumen - eficiencia
            Case 15, 16, 17 ' x producto,x personal,x tarea
                .Rows = 2
                .FixedRows = 2
                .FrozenCols = 2
                
                UNIR_CELDAS Grid, 0, 1, 0, 5, "-", flexAlignCenterCenter
                UNIR_CELDAS Grid, 0, 6, 0, 8, "Unid. x Hora", flexAlignCenterCenter
                UNIR_CELDAS Grid, 0, 9, 0, 11, "Eficiencias", flexAlignCenterCenter
                
                .TextMatrix(1, 1) = "Tarea":        .ColWidth(1) = 3500:   .ColAlignment(1) = flexAlignLeftBottom:          .Row = 1: .Col = 1: .CellAlignment = flexAlignLeftBottom
                .TextMatrix(1, 2) = "Producto":     .ColWidth(2) = 3500:   .ColAlignment(2) = flexAlignLeftBottom:          .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftBottom
                .TextMatrix(1, 3) = "Personal":     .ColWidth(3) = 0:      .ColAlignment(3) = flexAlignLeftBottom:          .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftBottom
                .TextMatrix(1, 4) = "Total":        .ColWidth(4) = 1000:   .ColAlignment(4) = flexAlignRightCenter:         .Row = 1: .Col = 4: .CellAlignment = flexAlignRightCenter
                .TextMatrix(1, 5) = "U.M.":         .ColWidth(5) = 450:    .ColAlignment(5) = flexAlignCenterCenter:        .Row = 1: .Col = 5: .CellAlignment = flexAlignCenterCenter
                ' und. x hora
                .TextMatrix(1, 6) = "Mínimo":      .ColWidth(6) = 870:     .ColAlignment(6) = flexAlignRightCenter:         .Row = 1: .Col = 6: .CellAlignment = flexAlignRightCenter
                .TextMatrix(1, 7) = "Promedio":    .ColWidth(7) = 870:     .ColAlignment(7) = flexAlignRightCenter:         .Row = 1: .Col = 7: .CellAlignment = flexAlignRightCenter
                .TextMatrix(1, 8) = "Máximo":      .ColWidth(8) = 870:     .ColAlignment(8) = flexAlignRightCenter:         .Row = 1: .Col = 8: .CellAlignment = flexAlignRightCenter
                ' eficiencia
                .TextMatrix(1, 9) = "Mínimo":      .ColWidth(9) = 710:     .ColAlignment(9) = flexAlignRightCenter:         .Row = 1: .Col = 9: .CellAlignment = flexAlignRightCenter
                .TextMatrix(1, 10) = "Promedio":   .ColWidth(10) = 710:    .ColAlignment(10) = flexAlignRightCenter:        .Row = 1: .Col = 10: .CellAlignment = flexAlignRightCenter
                .TextMatrix(1, 11) = "Máximo":     .ColWidth(11) = 710:    .ColAlignment(11) = flexAlignRightCenter:        .Row = 1: .Col = 11: .CellAlignment = flexAlignRightCenter
                .TextMatrix(1, 12) = "Cant. Reg":  .ColWidth(12) = 800:    .ColAlignment(12) = flexAlignRightCenter:        .Row = 1: .Col = 12: .CellAlignment = flexAlignRightCenter
                                
                If mTipoConsulta = 15 Then .ColWidth(2) = 0 ' agrupado por producto
                If mTipoConsulta = 17 Then .ColWidth(1) = 0 ' agrupado por tarea
        End Select

        ' DE LOS ID'S
        For k = 1 To Q_COL_FILA_OCULTA
            .TextMatrix(0, k) = "ID" + CStr(k):         .ColWidth(k) = 500
        Next
        If Q_COL_FILA_OCULTA <> -1 Then OCULTAR_COL Grid, 1, Q_COL_FILA_OCULTA
    End With
    DoEvents
End Sub

'*****************************************************************************************************
'* Nombre           : PONER_FORMATO
'* Tipo             : FUNCION
'* Descripcion      : ESTA FUNCION CONVERTIRA AL FORMATO
'* Paranetros       : NOMBRE          |  TIPO        |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    S_MONTO         |  Double      |
'*                    Band_Total_gral |  Boolean     |
'*                    Q_POS           |  Integer     |
'* Devuelve         :
'*****************************************************************************************************
Private Function PONER_FORMATO(S_MONTO As Double, Optional Band_Total_gral As Boolean = False, Optional Q_POS As Integer = -1) As String
    If S_MONTO = 0 Then
            PONER_FORMATO = "0.00"
        Exit Function
    End If
    
    PONER_FORMATO = Format(S_MONTO, FORMAT_MONTO)
End Function

Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    Dim xCampos() As String
    Dim nTitulo As String
    Dim nSQL As String
    Dim nSQLNotIn As String
    Dim Q_ROW As Long
    
    If Col <> 2 Then Exit Sub
    Select Case Index
        Case 0 ' personal
            If NulosC(fg(Index).TextMatrix(Row, Col)) <> "" Then
                nSQLNotIn = vbCr + " WHERE UCASE([pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom]) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%'"
            End If
        
            ReDim xCampos(3, 4) As String
            xCampos(0, 0) = "Apellidos y Nombres":  xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "4500":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
            xCampos(1, 0) = "Fch. Nac":             xCampos(1, 1) = "fchnac":       xCampos(1, 2) = "1000":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
            xCampos(2, 0) = "Id":                   xCampos(2, 1) = "id":           xCampos(2, 2) = "700":      xCampos(2, 3) = "C":    xCampos(1, 4) = "N"
            
            nSQL = "SELECT pla_empleados.id , [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS descripcion, pla_empleados.fchnac " _
                + vbCr + " FROM pla_empleados " & nSQLNotIn _
                + vbCr + " ORDER BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom]; "

            nTitulo = "Buscando Personal"
            
        Case 1 ' producto
            ReDim xCampos(3, 4) As String
            xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "codpro":       xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
            xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "5000":    xCampos(1, 3) = "C"
            xCampos(2, 0) = "id":           xCampos(2, 1) = "id":           xCampos(2, 2) = "700":     xCampos(2, 3) = "N"
                          
            ' de los registros ya seleccionados
            nSQLNotIn = GENERAR_SQL_ID(fg(Index), 1, " and alm_inventario.id", "NOT IN")
            
            If NulosC(fg(Index).TextMatrix(Row, Col)) <> "" Then
                nSQLNotIn = nSQLNotIn & " AND UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' "
            End If
            
            nSQL = "SELECT alm_inventario.id,alm_inventario.codpro, alm_inventario.descripcion " _
                + vbCr + " FROM alm_inventario " _
                + vbCr + " WHERE alm_inventario.id IN (SELECT pro_receta.iditem FROM pro_receta INNER JOIN pro_recetatar ON pro_receta.id = pro_recetatar.idrec;) " & nSQLNotIn _
                + vbCr + " ORDER BY alm_inventario.descripcion; "
            
            nTitulo = "Buscando Producto"
            
        Case 2 ' tarea
            ReDim xCampos(4, 4) As String
            xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "descripcion":     xCampos(0, 2) = "4500":    xCampos(0, 3) = "C"
            xCampos(1, 0) = "Abrev":        xCampos(1, 1) = "nomcorto":   xCampos(1, 2) = "2300":    xCampos(1, 3) = "C"
            xCampos(2, 0) = "Diverso":      xCampos(2, 1) = "diverso":    xCampos(2, 2) = "700":     xCampos(2, 3) = "C"
            xCampos(3, 0) = "Id":           xCampos(3, 1) = "id":         xCampos(3, 2) = "600":     xCampos(3, 3) = "N"
            
            ' de los registros ya seleccionados
            nSQLNotIn = GENERAR_SQL_ID(fg(Index), 1, " WHERE pro_tareas.id", "NOT IN")
            
            If NulosC(fg(Index).TextMatrix(Row, Col)) <> "" Then
                If nSQLNotIn = "" Then
                    nSQLNotIn = " WHERE "
                Else
                    nSQLNotIn = nSQLNotIn & " AND "
                End If
                
                nSQLNotIn = nSQLNotIn & " (UCASE(pro_tareas.descripcion) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' OR UCASE(pro_tareas.abrev) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' ) "
            End If
            
            nSQL = "SELECT pro_tareas.id, pro_tareas.codigo, pro_tareas.descripcion , pro_tareas.abrev AS nomcorto, mae_unidades.id AS idunimed, mae_unidades.abrev, IIf([pro_tareas].[diverso]=-1,'Si','No') AS diverso " _
                    + vbCr + " FROM mae_unidades RIGHT JOIN pro_tareas ON mae_unidades.id = pro_tareas.idunimed  " & nSQLNotIn
            nTitulo = "Buscando Tareas"
        
        Case Else
            Exit Sub
    End Select

    Dim xRs As New ADODB.Recordset
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    
    fg(Index).TextMatrix(fg(Index).Row, 1) = NulosC(xRs.Fields("id"))
    fg(Index).TextMatrix(fg(Index).Row, 2) = NulosC(xRs.Fields("descripcion"))
        
    If fg(Index).Row = fg(Index).Rows - 1 Then fg(Index).AddItem ""
    fg(Index).Row = fg(Index).Rows - 1: fg(Index).Col = 2
        
SALIR:
    Set xRs = Nothing
    Exit Sub

error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Fg_CellButtonClick(" + CStr(Index) + ")", True, "Error..."
End Sub

Private Sub fg_DblClick(Index As Integer)
    Fg_CellButtonClick Index, fg(Index).Rows - 1, 2
End Sub

Private Sub Fg_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If fg(Index).Row = -2 Then Exit Sub
    Select Case KeyCode
        Case 45  ' INSERTAR REGI
            fg(Index).AddItem ""
            fg(Index).Row = fg(Index).Rows - 1: fg(Index).Col = 2
        
        Case 46  ' SUPRIMIR/DELETE
            If fg(Index).Rows - 1 >= 2 Then
                fg(Index).RemoveItem fg(Index).Row
                fg(Index).Row = fg(Index).Rows - 1: fg(Index).Col = 2
            Else
                LimpiarGrid fg(Index), True
                GRID_COMBOLIST fg(Index)
            End If
    End Select
End Sub

'*****************************************************************************************************
'* Nombre           : fEstiloConsulta
'* Tipo             : FUNCION
'* Descripcion      : Establecer la estructura de la consulta, EL VALOR QUE DEVUELVE definira la
'*                    estructura del grid
'* Paranetros       : NOMBRE    |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    mTipo     |  Integer    |  :Detalle, 2::Resumen
'* Devuelve         : Integer
'*****************************************************************************************************
Private Function fEstiloConsulta(mTipo As Integer) As Integer
    Dim mTipoEstilo As Integer
    If mTipo = 1 Then     ' detalle
        If OptConsulta(0).Value = True Then                   ' control de tarea
            If OptGrupo(0).Value = True Then mTipoEstilo = 0  ' X Producto
            If OptGrupo(1).Value = True Then mTipoEstilo = 1  ' X Personal
            If OptGrupo(2).Value = True Then mTipoEstilo = 2  ' X Tarea
        ElseIf OptConsulta(1).Value = True Then               ' eficiencia
            If OptGrupo(0).Value = True Then mTipoEstilo = 3  ' X Producto
            If OptGrupo(1).Value = True Then mTipoEstilo = 4  ' X Personal
            If OptGrupo(2).Value = True Then mTipoEstilo = 5  ' X Tarea
        Else              ' minimos y maximos
            If OptGrupo(0).Value = True Then mTipoEstilo = 6  ' X Producto
            If OptGrupo(1).Value = True Then mTipoEstilo = 7  ' X Personal
            If OptGrupo(2).Value = True Then mTipoEstilo = 8  ' X Tarea
        End If
    Else                  ' resumen
        If OptConsulta(0).Value = True Then                   ' control de tarea
            If OptGrupo(0).Value = True Then mTipoEstilo = 9  ' X Producto
            If OptGrupo(1).Value = True Then mTipoEstilo = 10 ' X Personal
            If OptGrupo(2).Value = True Then mTipoEstilo = 11 ' X Tarea
        ElseIf OptConsulta(1).Value = True Then               ' eficiencia
            If OptGrupo(0).Value = True Then mTipoEstilo = 12 ' X Producto
            If OptGrupo(1).Value = True Then mTipoEstilo = 13 ' X Personal
            If OptGrupo(2).Value = True Then mTipoEstilo = 14 ' X Tarea
        Else              ' minimos y maximos
            If OptGrupo(0).Value = True Then mTipoEstilo = 15 ' X Producto
            If OptGrupo(1).Value = True Then mTipoEstilo = 16 ' X Personal
            If OptGrupo(2).Value = True Then mTipoEstilo = 17 ' X Tarea
        End If
    End If
    
    fEstiloConsulta = mTipoEstilo
End Function

'*****************************************************************************************************
'* Nombre           : PosicionarProgBar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : POSICIONAR EL PROGRESO DENTRO DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub PosicionarProgBar()
    FraProgreso.Left = (Me.Width - FraProgreso.Width) \ 2
    FraProgreso.Top = (Me.Height - FraProgreso.Height) \ 2
    Me.PgBar.Value = 1
    FraProgreso.Visible = True
End Sub

Private Sub Fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Index = 0 Then PopupMenu Menu1
        If Index = 1 Then PopupMenu Menu2
    End If
End Sub

' DEL PRODUCTO
Private Sub Menu1_1_Click()
    fg_DblClick 0
End Sub

Private Sub Menu1_3_Click()
    Fg_KeyDown 0, 46, 0
End Sub

' DE LOS INSUMOS
Private Sub Menu2_1_Click()
    fg_DblClick 1
End Sub

Private Sub Menu2_3_Click()
    Fg_KeyDown 1, 46, 0
End Sub

'*****************************************************************************************************
'* Nombre           : pExportarExcel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA A MS EXCEL LOS DATOS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pExportarExcel()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    X_PRINT.VSFlexGrid_Exportar_MSExcel xCon, IIf(TabOne1.CurrTab = 0, Fg1, Fg2), T_RPT_TITULO + " ", T_RPT_PERIODO, , "Registro de Tareas"
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
    
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportarExcel"
End Sub

Private Sub OptConsulta_Click(Index As Integer)
    habilitar OptGrupo, True
    OptGrupo(0).Value = True
   
    If Index = 0 Then          ' control tarea
        Fra_Top.Enabled = False
        OptGrupo(2).Enabled = True
    ElseIf Index = 1 Then      ' eficiencia de personal
        Fra_Top.Enabled = True
        OptGrupo(2).Enabled = True
    Else                       ' maximos y minimos
        Fra_Top.Enabled = False
        txt_top.Text = ""
    End If
End Sub

Private Sub pic_Click(Index As Integer)
    pHabilitarBotonEditor False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 5 Then pExportarExcel
    If Button.Index = 6 Then pImprimir
    If Button.Index = 8 Then
        BAND_INTERRUMPIR = True
        Unload Me
    End If
End Sub

Private Sub cb_Click(Index As Integer)
    Dim xCampos() As String
    Dim nCampoBusca As String
    Dim nSQL As String
    Dim nTitulo As String
    On Error GoTo error
    Select Case Index
        Case 0 ' area
            nTitulo = "Buscando Area"
            nSQL = "SELECT mae_area.id, mae_area.descripcion AS nombre, mae_area.id AS cod, mae_area.id AS idarea " _
                + vbCr + " FROM pro_area INNER JOIN mae_area ON pro_area.idarea = mae_area.id; "
    End Select
    
    ReDim xCampos(2, 3) As String
    xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id":           xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
    
    Dim RstTmp As New ADODB.Recordset
    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio

    If RstTmp.State = 0 Then GoTo SALIR
    If RstTmp.EOF = True Or RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index).Text = NulosC(RstTmp.Fields(0))         ' TEXTO A MOSTRAR
    lbl_cb(Index).Caption = NulosC(RstTmp.Fields(1))      ' NOMBRE
    lbl_cod(Index).Caption = NulosN(RstTmp.Fields(2))     ' CODIGO
    lbl_cb(Index).ToolTipText = NulosC(RstTmp.Fields(1))  ' NOMBRE
      
    Select Case Index
        Case 0
            fg(0).SetFocus
    End Select

SALIR:
    Set RstTmp = Nothing
    Exit Sub

error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cod(Index).Caption = ""
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
        If Index <> 1 Then
            SendKeys vbTab
        Else
            If Fg1.Rows >= 2 Then
                Fg1.Row = 1: Fg1.Col = 1
            Else
                Fg1.Row = Fg1.Rows - 1: Fg1.Col = 1
            End If
            Fg1.SetFocus
        End If
        Exit Sub
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)
    If txt_cb(Index).Text = "" Then Exit Sub
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
        Case 0 ' area
            nSQL = "SELECT mae_area.id, mae_area.descripcion AS nombre, mae_area.id AS cod, mae_area.id AS idarea " _
                + vbCr + " FROM pro_area INNER JOIN mae_area ON pro_area.idarea = mae_area.id; "
        
        Case Else
            Exit Sub
    End Select

    If xCon.State = 0 Then GoTo SALIR
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo SALIR

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cb(Index).Text = NulosC(RstTmp.Fields(0))        ' TEXTO A MOSTRAR
        lbl_cb(Index).Caption = NulosC(RstTmp.Fields(1))     ' NOMBRE
        lbl_cod(Index).Caption = NulosN(RstTmp.Fields(2))    ' CODIGO
        lbl_cb(Index).ToolTipText = NulosC(RstTmp.Fields(1)) ' NOMBRE
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cod(Index).Caption = ""
    End If
    
    Set RstTmp = Nothing
    Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "txt_cb_Validate (" + CStr(Index) + ")"
    Exit Sub
    
SALIR:
    Set RstTmp = Nothing
    txt_cb(Index).Text = ""
End Sub

Private Sub txt_top_Change()
    If NulosN(txt_top.Text) = 0 Then
        cb_top.ListIndex = 0
    End If
End Sub

Private Sub txt_top_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarDatosGrid
'* Tipo             : FUNCION
'* Descripcion      : FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
'* Paranetros       : NOMBRE      |  TIPO            |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Grid        |  VSFlexGrid      |
'*                    RST_ORIGEN  |  ADODB.Recordset |
'* Devuelve         :
'*****************************************************************************************************
Private Function pCargarDatosGrid(Grid As VSFlexGrid, RST_ORIGEN As ADODB.Recordset)
    Dim BAND_ADD_REG As Boolean
    
    BAND_ADD_REG = True
    
    If RST_ORIGEN.RecordCount > 0 Then
        RST_ORIGEN.MoveFirst
    Else
        Exit Function
    End If
    PgBar.Min = 0
    PgBar.Max = RST_ORIGEN.RecordCount
    
    While Not RST_ORIGEN.EOF
    
    DoEvents
        ' SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then Exit Function
        pCompararGrupo Grid, RST_ORIGEN, BAND_ADD_REG, Q_COL_COMPARAR_GRUPO
        If RST_ORIGEN.Bookmark <> 1 Then ADD_REG Grid
        ' CARGAR A LA GRILLA
        pCargarDatosGridArrayTmp Grid, RST_ORIGEN, Grid.Rows - 1
        ' PONER COLOR FONDO
        If Q_COL_COMPARAR_FONDO <> -1 Then pCargarDatosGridFondo Grid, RST_ORIGEN, Grid.Rows - 1, 1, Grid.Rows - 1, Grid.Cols - 1
        
        RST_ORIGEN.MoveNext

        ' PONER TOTALES AL FINAL DE LA GRILLA
        If RST_ORIGEN.EOF Then
            pCargarDatosGridAddTotales Grid, BAND_ADD_REG, "Total:"
            Select Case ESTILO_VISTA
                Case 0, 1, 2, 4, 5, 8, 9
                
                Case Else
                    pCargarDatosGridAddTotales Grid, True, "Tot Gen:", True
            
            End Select
        Else
            PgBar.Value = CLng(RST_ORIGEN.Bookmark)
        End If
    Wend
End Function

'*****************************************************************************************************
'* Nombre           : pCompararGrupo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PROCEDIMIENTO QUE NOS PERMITE ARMAR LOS GRUPOS, OMPARA CUANDO CAMBIAR DE GRUPO
'* Paranetros       : NOMBRE           |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Grid             |  VSFlexGrid       |
'*                    RST_ORIGEN       |  ADODB.Recordset  |
'*                    BAND_ADD_REG     |  Boolean          |
'*                    Q_COL_COMPARAR   |  Integer          |
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCompararGrupo(Grid As VSFlexGrid, RST_ORIGEN As ADODB.Recordset, BAND_ADD_REG As Boolean, Optional Q_COL_COMPARAR As Integer = -1)
    Dim RST_TEPM_1 As New ADODB.Recordset
    Dim N_GRUPO_ADD As String
    Dim Q_POS As Integer
    
    If Q_COL_COMPARAR = -1 Then
        If RST_ORIGEN.Bookmark = 1 Then ADD_REG Grid, Fila_Ninguno
        Exit Sub
    End If
    
    If Q_COL_COMPARAR = -1 Then Q_COL_COMPARAR = Q_COL_COMPARAR_GRUPO
    
    If Q_COL_GRUPO_ADD <> -1 Then
        If NulosC(N_CAMPO_GRUPO_ADD) <> "" Then
            For Q_POS = 1 To Q_COL_GRUPO_ADD
                If LCase(RST_ORIGEN.Fields(Q_COL_COMPARAR + Q_POS).Name) = UCase(N_CAMPO_GRUPO_ADD) Then
                    N_GRUPO_ADD = Format(NulosN(RST_ORIGEN.Fields(Q_COL_COMPARAR + Q_POS)), FORMAT_MONTO) + " " + N_GRUPO_ADD
                Else
                    N_GRUPO_ADD = NulosC(RST_ORIGEN.Fields(Q_COL_COMPARAR + Q_POS)) & "  " + N_GRUPO_ADD
                End If
            Next Q_POS
        End If
        N_GRUPO_ADD = "  =>>   " + N_GRUPO_ADD
    End If
    
    If RST_ORIGEN.Bookmark = 1 Then
        ' SE CARGA EN fGenerarConsulta() Q_COL_COMPARAR_GRUPO
        ADD_REG Grid, Fila_grupo
        UNIR_CELDAS Grid, Grid.Rows - 1, Q_COL_GRUPO_INICIO, Grid.Rows - 1, Q_COL_GRUPO_TERMINA, INICIO_GRUPO & NulosC(RST_ORIGEN.Fields(Q_COL_COMPARAR)) & N_GRUPO_ADD, flexAlignLeftCenter:
        FORMATO_CELDA Grid, Grid.Rows - 1, Q_COL_GRUPO_INICIO
        ADD_REG Grid, Fila_Ninguno
        UNIR_CELDAS Grid, Grid.Rows - 1, 1, Grid.Rows - 1, Grid.Cols - 1, " ", flexAlignLeftCenter
        nSQLValor_FONDO = ""
    Else
        Set RST_TEPM_1 = RST_ORIGEN.Clone
        RST_TEPM_1.Bookmark = RST_ORIGEN.Bookmark
        RST_TEPM_1.MovePrevious
        
        If NulosC(RST_TEPM_1.Fields(Q_COL_COMPARAR)) <> NulosC(RST_ORIGEN.Fields(Q_COL_COMPARAR)) Then
            ' cargar datos de total
            pCargarDatosGridAddTotales Grid, BAND_ADD_REG, "Total:"
            
            ' poner la fila en blanco, agrupado
            ADD_REG Grid, Fila_en_Blanco
            UNIR_CELDAS Grid, Grid.Rows - 1, IIf(Q_COL_FILA_OCULTA = -1, 1, Q_COL_FILA_OCULTA + 1), Grid.Rows - 1, Grid.Cols - 1, " ", flexAlignLeftCenter
            Limpiar_ARRAY_TOTAL
            ADD_REG Grid, Fila_grupo
            UNIR_CELDAS Grid, Grid.Rows - 1, Q_COL_GRUPO_INICIO, Grid.Rows - 1, Q_COL_GRUPO_TERMINA, INICIO_GRUPO & NulosC(RST_ORIGEN.Fields(Q_COL_COMPARAR)) & N_GRUPO_ADD, flexAlignLeftCenter
            FORMATO_CELDA Grid, Grid.Rows - 1, Q_COL_GRUPO_INICIO
            
            ' inicializando el color del fondo
            nSQLValor_FONDO = ""
        End If
    End If

SALIR:
    Set RST_TEPM_1 = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarDatosGridArrayTmp
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : FUNCION QUE AGREGARA LOS REGISTROS AL CONTRO FLEGRID
'* Paranetros       : NOMBRE      |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Grid        |  VSFlexGrid       |
'*                    RST_ORIGEN  |  ADODB.Recordset  |
'*                    Q_ROW       |  Long             |
'* Devuelve         :
'*****************************************************************************************************
Private Function pCargarDatosGridArrayTmp(Grid As VSFlexGrid, RST_ORIGEN As ADODB.Recordset, Q_ROW As Long)
    Dim Q_INCREMENTO_X_COL As Integer   ' SERVIRA PARA POSICIONAR LA PRIMERA COLUMNA DE ENERO DE UN AÑO
    Dim Q_POS_MES_TOTAL  As Integer     ' POSICIONA A LA COLUMNA DEL TOTAL X AÑO
    Dim Q_POS As Integer
    Dim Q_CAMPO As Integer
    Dim vStrCampo As String
    
    ' IDENTIFICAR LA POSICION DE INICIO DE MES(ENERO)
    Q_INCREMENTO_X_COL = 0
    Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
    
    DoEvents
    
    For Q_CAMPO = 0 To RST_ORIGEN.Fields.Count - 1
        If BAND_INTERRUMPIR = True Then Exit Function
        vStrCampo = RST_ORIGEN.Fields(Q_CAMPO).Name
        
        Select Case LCase(vStrCampo)
            Case "acumulado", "total"
                Grid.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO)
            
            Case "horini", "horfin"
                Grid.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(RST_ORIGEN.Fields(vStrCampo), FORMAT_HORA_SIN_SEGUNDO)
            
            Case "fchtra"
                Grid.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(RST_ORIGEN.Fields(vStrCampo), FORMAT_DATE)
                
            Case "eficiencia", "eficmin", "eficprom", "eficmax"
                If NulosN(RST_ORIGEN.Fields(vStrCampo)) = 100 Then         ' negro (consumo ahorrado)
                    Grid.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(RST_ORIGEN.Fields(vStrCampo), FORMAT_PORCENTAJE) & "%"
                ElseIf NulosN(RST_ORIGEN.Fields(vStrCampo)) = 0 Then       ' no mostrar datos
                    
                ElseIf NulosN(RST_ORIGEN.Fields(vStrCampo)) > 100 Then     ' azul (supera la eficiencia)
                    FORMATO_CELDA Grid, Q_ROW, Q_CAMPO + 1, &HFF0000, False, &HFFFFFF, Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_PORCENTAJE) + "%"
                ElseIf NulosN(RST_ORIGEN.Fields(vStrCampo)) < 100 Then     ' rojo (menos eficiente)
                    FORMATO_CELDA Grid, Q_ROW, Q_CAMPO + 1, &HFF, False, &HFFFFFF, Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_PORCENTAJE) + "%"
                End If
                
            Case "unidxmin", "unidxhor"
                Grid.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), "#,##0.00000")
            
            Case "canhrmin", "canhrprom", "canhrmax"
                Grid.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), "#,##0.000")
                
            Case "totmin"       ' total minutos
                Grid.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(RST_ORIGEN.Fields(vStrCampo), FORMAT_MONTO)
            Case "difhor"
                Grid.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(RST_ORIGEN.Fields(vStrCampo), FORMAT_HORA_LARGO)
            Case "cantreal", "canteo"
                Grid.TextMatrix(Q_ROW, Q_CAMPO + 1) = Format(RST_ORIGEN.Fields(vStrCampo), FORMAT_MONTO)
            
            Case Else
                ' AGREGAR LOS DEMAS DATOS
                Grid.TextMatrix(Q_ROW, Q_CAMPO + 1) = RST_ORIGEN.Fields(vStrCampo) & ""
        End Select
    Next
End Function

'*****************************************************************************************************
'* Nombre           : pCargarDatosGridFondo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PONER COLOR FONDO
'* Paranetros       : NOMBRE      |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Grid        |  VSFlexGrid       |
'*                    RST_ORIGEN  |  ADODB.Recordset  |
'*                    X_ROW1      |  Long             |
'*                    X_COL1      |  Integer          |
'*                    X_ROW2      |  Long             |
'*                    X_COL2      |  Integer          |
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarDatosGridFondo(Grid As VSFlexGrid, RST_ORIGEN As ADODB.Recordset, X_ROW1 As Long, X_COL1 As Integer, X_ROW2 As Long, X_COL2 As Integer)
    If Q_COL_COMPARAR_FONDO = -1 Then Exit Sub
        If NulosN(Grid.TextMatrix(X_ROW1, 1)) = e_ESTADO_ROW_GRID.Fila_grupo Then
        ' SI SE DESEA PONER COLOR AL GRUPO
        ElseIf NulosN(Grid.TextMatrix(X_ROW1, 1)) = e_ESTADO_ROW_GRID.Fila_Total Then
        ElseIf NulosN(Grid.TextMatrix(X_ROW1, 1)) = e_ESTADO_ROW_GRID.Fila_Total_grl Then
        ElseIf NulosN(Grid.TextMatrix(X_ROW1, 1)) = e_ESTADO_ROW_GRID.Fila_en_Blanco Then
        Else
           If nSQLValor_FONDO = "" Then
                ' se coloca la opcion "-" para considerar los nulos
                nSQLValor_FONDO = NulosC(RST_ORIGEN.Fields(Q_COL_COMPARAR_FONDO)) & "-"
                nSQLValor_FONDO_COLOR = &HFDFFFF
                F_CAMIAR_FONDO = False
            End If
    
            If nSQLValor_FONDO = NulosC(RST_ORIGEN.Fields(Q_COL_COMPARAR_FONDO)) & "-" Then
                nSQLValor_FONDO_COLOR = nSQLValor_FONDO_COLOR
            Else
                nSQLValor_FONDO = NulosC(RST_ORIGEN.Fields(Q_COL_COMPARAR_FONDO)) & "-"
                If F_CAMIAR_FONDO = True Then
                    nSQLValor_FONDO_COLOR = &HFDFFFF
                    F_CAMIAR_FONDO = False
                Else
                    nSQLValor_FONDO_COLOR = &HE0FEFE
                    F_CAMIAR_FONDO = True
                End If
            End If
            GRID_COLOR_FONDO Grid, X_ROW1, X_COL1, X_ROW2, X_COL2, nSQLValor_FONDO_COLOR
        End If
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarDatosGridAddTotales
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : AGREGA LOS TOTALES POR CADA GRUPO Y EL TOTAL GENERAL, ACUMULA LOS TOTALES EN EL TOTAL GENERAL
'* Paranetros       : NOMBRE           |  TIPO         |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Grid             |  VSFlexGrid   |
'*                    BAND_ADD_TOTAL   |  Boolean      |
'*                    Nombre_total     |  String       |
'*                    Band_Total_gral  |  Boolean      |
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarDatosGridAddTotales(Grid As VSFlexGrid, BAND_ADD_TOTAL As Boolean, Nombre_total As String, Optional Band_Total_gral As Boolean = False)
    Dim Q_MES As Integer
    Dim X_ROW As Integer
''''''    'On Error Resume Next
''''''    X_ROW = Grid.Rows
''''''    If BAND_ADD_TOTAL = True Then
''''''        '--AGREAGNDO NUEVA FILA
''''''        ADD_REG Grid, IIf(Band_Total_gral = False, Fila_Total, Fila_Total_grl)
''''''
''''''        'PONIENDO LOS NOMBRES DE LOS TOTALES  Q_POSICION_TOTAL SE OBTIENE DE fGenerarConsulta()
''''''        Grid.TextMatrix(X_ROW, Q_POSICION_TOTAL) = Nombre_total
''''''        FORMATO_CELDA Grid, X_ROW, Q_POSICION_TOTAL
''''''    End If
''''''
''''''
''''''    '--ACUMULANDO LOS TOTALES GRLES
''''''    If Band_Total_gral = False Then
''''''        For Q_MES = 0 To Q_COL_ARR_TOTAL
''''''            ARR_TMP(Q_MES, 4) = NulosN(ARR_TMP(Q_MES, 4)) + NulosN(ARR_TMP(Q_MES, 3))
''''''        Next Q_MES
''''''        If Q_COL_FILA_ULTIMO <> -1 Then
''''''            ARR_TMP_1(0, 1) = NulosN(ARR_TMP_1(0, 1)) + NulosN(ARR_TMP_1(0, 0)) '--STOCK
''''''            ARR_TMP_1(1, 1) = NulosN(ARR_TMP_1(1, 1)) + NulosN(ARR_TMP_1(1, 0)) '--SALDO
''''''        End If
''''''    End If
''''''    '
'''''''--------------------------
''''''    Dim Q_INCREMENTO_X_COL As Integer   '--SERVIRA PARA POSICIONAR LA PRIMERA COLUMNA DE ENERO DE UN AÑO
''''''    Dim Q_POS_MES_TOTAL  As Integer     '--POSICIONA A LA COLUMNA DEL TOTAL X AÑO
''''''
''''''
''''''    '--IDENTIFICAR LA POSICION DE INICIO DE MES(ENERO)
''''''    Q_INCREMENTO_X_COL = 0
''''''    Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
''''''    '-----------
''''''
''''''    Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
''''''
''''''    For Q_MES = 0 To Q_COL_ARR_TOTAL
''''''        '--INTERRUMPIR EL PROCESO
''''''        If BAND_INTERRUMPIR = True Then Exit Sub
''''''        Grid.TextMatrix(X_ROW, Q_POS_MES) = PONER_FORMATO(IIf(Band_Total_gral = False, ARR_TMP(Q_MES, 3), ARR_TMP(Q_MES, 4)), Band_Total_gral, Q_MES)
''''''        FORMATO_CELDA Grid, X_ROW, Q_POS_MES
''''''        Q_POS_MES = Q_POS_MES + 1
''''''    Next Q_MES
''''''
''''''
''''''    If Q_COL_FILA_ULTIMO <> -1 Then
''''''        '--STOCK
''''''        Grid.TextMatrix(X_ROW, Grid.Cols - 2) = PONER_FORMATO(IIf(Band_Total_gral = False, ARR_TMP_1(0, 0), ARR_TMP_1(0, 1)), Band_Total_gral, Grid.Cols - 2)
''''''        FORMATO_CELDA Grid, X_ROW, Grid.Cols - 2, RGB(128, 0, 0)
''''''        '--SALDO
''''''        Grid.TextMatrix(X_ROW, Grid.Cols - 1) = PONER_FORMATO(IIf(Band_Total_gral = False, ARR_TMP_1(1, 0), ARR_TMP_1(1, 1)), Band_Total_gral, Grid.Cols - 1)
''''''        FORMATO_CELDA Grid, X_ROW, Grid.Cols - 1, vbRed
''''''    End If
''''''    Err.Clear
End Sub

'*****************************************************************************************************
'* Nombre           : pHabilitarBotonEditor
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : Mostrar el Cuadro de Opciones, Mostrar/Ocultar las opciones
'* Paranetros       : NOMBRE    |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    band      |  Boolean    |  true (mostrar el cuadro), false (ocultar el cuadro)
'* Devuelve         :
'*****************************************************************************************************
Private Sub pHabilitarBotonEditor(band As Boolean)
    Dim A&
    ' si es true cargar los datos
    CmdOpcion.Enabled = Not band
    
    TabOne1.TabEnabled(0) = Not band
    TabOne1.TabEnabled(1) = Not band
    
    For A = 1 To Toolbar1.Buttons.Count - 1
        Toolbar1.Buttons(A).Enabled = Not band
    Next A
    
    habilitar TxtFecha, Not band
    habilitar OptConsulta, Not band
    habilitar OptGrupo, Not band
    habilitar fg, Not band
    
    If txt_top.Visible = True Then
        habilitar cb, Not band
        habilitar txt_cb, Not band
        
        txt_top.Enabled = Not band
        cb_top.Enabled = Not band
        
    End If
    
     FraOpcion.Visible = band
     
    ' true muestra el ingreso de datos
    If band = True Then
        FraOpcion.Top = 2820
        FraOpcion.Left = 3810
    End If
End Sub

Private Sub txtObs_KeyPress(KeyAscii As Integer)
    If validar_letras(KeyAscii) = False Then
        If validar_numero(KeyAscii) = False Then
            KeyAscii = 0
        End If
    End If
End Sub
