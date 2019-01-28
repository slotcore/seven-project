VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmControlTarea1 
   Caption         =   "Producción - Ingreso de Tareas"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   -1770
   ClientWidth     =   19080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   19080
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame FraEditor 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   11250
      TabIndex        =   100
      Top             =   8130
      Visible         =   0   'False
      Width           =   5205
      Begin VB.CommandButton CmdEditor 
         Caption         =   "&Salir"
         Height          =   420
         Index           =   2
         Left            =   4200
         TabIndex        =   104
         Top             =   2415
         Width           =   945
      End
      Begin VB.CommandButton CmdEditor 
         Caption         =   "Eliminar"
         Height          =   420
         Index           =   1
         Left            =   4200
         TabIndex        =   103
         Top             =   1815
         Width           =   945
      End
      Begin VB.CommandButton CmdEditor 
         Caption         =   "Agregar"
         Height          =   420
         Index           =   0
         Left            =   4200
         TabIndex        =   102
         Top             =   1320
         Width           =   945
      End
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   4935
         Picture         =   "FrmControlTarea1.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   101
         ToolTipText     =   "Cerrar"
         Top             =   30
         Width           =   195
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   2445
         Index           =   1
         Left            =   60
         TabIndex        =   105
         Top             =   360
         Width           =   4050
         _cx             =   7144
         _cy             =   4313
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
         Rows            =   12
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmControlTarea1.frx":02EC
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
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblTotal(3)"
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
         Index           =   3
         Left            =   4185
         TabIndex        =   110
         Top             =   975
         Width           =   870
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total P.N"
         Height          =   195
         Index           =   2
         Left            =   4185
         TabIndex        =   109
         Top             =   780
         Width           =   675
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblTotal(1)"
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
         Index           =   1
         Left            =   4185
         TabIndex        =   108
         Top             =   540
         Width           =   870
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total P.B"
         Height          =   195
         Index           =   0
         Left            =   4185
         TabIndex        =   107
         Top             =   360
         Width           =   660
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         Index           =   2
         X1              =   4140
         X2              =   4140
         Y1              =   345
         Y2              =   2800
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   0
         X1              =   30
         X2              =   5700
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -150
         X2              =   5895
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   1
         X1              =   5190
         X2              =   5190
         Y1              =   -120
         Y2              =   4770
      End
      Begin VB.Label LblTituloFrame 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Registros"
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
         TabIndex        =   106
         Top             =   60
         Width           =   810
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   15
         Y1              =   -210
         Y2              =   4845
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   270
         Index           =   1
         Left            =   30
         Top             =   15
         Width           =   5130
      End
   End
   Begin VB.Frame FraOpcion 
      BorderStyle     =   0  'None
      Height          =   2370
      Left            =   180
      TabIndex        =   88
      Top             =   7650
      Visible         =   0   'False
      Width           =   4020
      Begin VB.CommandButton CmdEditor 
         Caption         =   "Aceptar"
         Height          =   465
         Index           =   3
         Left            =   1275
         TabIndex        =   98
         Top             =   1830
         Width           =   1365
      End
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   3735
         Picture         =   "FrmControlTarea1.frx":03D6
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   97
         ToolTipText     =   "Cerrar"
         Top             =   30
         Width           =   195
      End
      Begin VB.CheckBox chkOpcion 
         Caption         =   "Quitar Filtro de Tarea por Area"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   96
         Top             =   390
         Value           =   1  'Checked
         Width           =   2430
      End
      Begin VB.Frame Frame5 
         Caption         =   "Al Insertar un Registro. >> Agregar..."
         Height          =   1125
         Left            =   105
         TabIndex        =   89
         Top             =   675
         Width           =   3840
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Nª Lote"
            Height          =   195
            Index           =   1
            Left            =   195
            TabIndex        =   95
            Top             =   255
            Value           =   1  'Checked
            Width           =   900
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Tipo"
            Height          =   195
            Index           =   2
            Left            =   195
            TabIndex        =   94
            Top             =   540
            Value           =   1  'Checked
            Width           =   750
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Personal / Nº Grupo"
            Height          =   195
            Index           =   3
            Left            =   195
            TabIndex        =   93
            Top             =   840
            Value           =   1  'Checked
            Width           =   1905
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Tarea"
            Height          =   195
            Index           =   4
            Left            =   2085
            TabIndex        =   92
            Top             =   255
            Value           =   1  'Checked
            Width           =   945
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Producto"
            Height          =   195
            Index           =   5
            Left            =   2085
            TabIndex        =   91
            Top             =   540
            Value           =   1  'Checked
            Width           =   1035
         End
         Begin VB.CheckBox chkOpcion 
            Caption         =   "Hora Inicio y Final"
            Height          =   195
            Index           =   6
            Left            =   2085
            TabIndex        =   90
            Top             =   840
            Width           =   1635
         End
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
         Y1              =   2340
         Y2              =   2355
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
      Begin VB.Label Label2 
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
         TabIndex        =   99
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
   Begin VB.Frame FraTarea 
      BorderStyle     =   0  'None
      Height          =   3720
      Left            =   4350
      TabIndex        =   75
      Top             =   7860
      Visible         =   0   'False
      Width           =   6450
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   2
         Left            =   6180
         Picture         =   "FrmControlTarea1.frx":06C2
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   81
         ToolTipText     =   "Cerrar"
         Top             =   30
         Width           =   195
      End
      Begin VB.CommandButton CmdTarea 
         Caption         =   "Cancelar"
         Height          =   420
         Index           =   0
         Left            =   2535
         TabIndex        =   80
         Top             =   3240
         Width           =   1365
      End
      Begin VB.CommandButton cb 
         Enabled         =   0   'False
         Height          =   240
         Index           =   3
         Left            =   1095
         Picture         =   "FrmControlTarea1.frx":09AE
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   675
         Width           =   225
      End
      Begin VB.CommandButton CmdTarea 
         Caption         =   "Agregar"
         Height          =   420
         Index           =   1
         Left            =   1050
         TabIndex        =   78
         Top             =   3240
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.OptionButton optTarea 
         Caption         =   "x Tarea"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   77
         Top             =   360
         Value           =   -1  'True
         Width           =   990
      End
      Begin VB.OptionButton optTarea 
         Caption         =   "x Receta"
         Height          =   195
         Index           =   1
         Left            =   1095
         TabIndex        =   76
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   3
         Left            =   555
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   83
         Text            =   "txt_cb(3)"
         Top             =   645
         Width           =   810
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   2205
         Index           =   0
         Left            =   45
         TabIndex        =   82
         Top             =   1005
         Width           =   6330
         _cx             =   11165
         _cy             =   3889
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
         Rows            =   12
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmControlTarea1.frx":0AE0
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
      Begin VB.Label lbl_cb 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cb(3)"
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
         Left            =   1380
         TabIndex        =   87
         Top             =   645
         Width           =   5010
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Buscar Tarea en Receta"
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
         TabIndex        =   86
         Top             =   60
         Width           =   2100
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   6435
         X2              =   6435
         Y1              =   -120
         Y2              =   4770
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
         X1              =   15
         X2              =   6435
         Y1              =   3690
         Y2              =   3690
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
      Begin VB.Label lbl_cod 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cod(3)"
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
         Index           =   3
         Left            =   3480
         TabIndex        =   85
         Top             =   645
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label lbl_cb_capt 
         AutoSize        =   -1  'True
         Caption         =   "Tarea"
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   84
         Top             =   750
         Width           =   420
      End
      Begin VB.Line Line9 
         X1              =   120
         X2              =   6285
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   270
         Index           =   2
         Left            =   30
         Top             =   15
         Width           =   6375
      End
   End
   Begin VB.Frame Frm4 
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
      Height          =   3600
      Left            =   11970
      TabIndex        =   48
      Top             =   3810
      Visible         =   0   'False
      Width           =   6120
      Begin VB.CheckBox chkAuto 
         Caption         =   "Cálculo A&utomático"
         Height          =   195
         Left            =   4260
         TabIndex        =   111
         Top             =   480
         Width           =   1695
      End
      Begin VB.Frame Frame8 
         Height          =   495
         Left            =   45
         TabIndex        =   53
         Top             =   3000
         Width           =   6015
         Begin VB.CommandButton CmdPer 
            Caption         =   "Agregar Grupo"
            Enabled         =   0   'False
            Height          =   330
            Index           =   4
            Left            =   2250
            TabIndex        =   64
            ToolTipText     =   "Agregar Personal"
            Top             =   135
            Width           =   1200
         End
         Begin VB.CommandButton CmdPer 
            Caption         =   "Elimi&nar"
            Enabled         =   0   'False
            Height          =   330
            Index           =   2
            Left            =   3660
            TabIndex        =   57
            TabStop         =   0   'False
            ToolTipText     =   "Eliminar Personal"
            Top             =   135
            Width           =   1035
         End
         Begin VB.CommandButton CmdPer 
            Caption         =   "Eliminar Todos"
            Enabled         =   0   'False
            Height          =   330
            Index           =   3
            Left            =   4695
            TabIndex        =   56
            ToolTipText     =   "Agregar Personal"
            Top             =   135
            Width           =   1200
         End
         Begin VB.CommandButton CmdPer 
            Caption         =   "&Agregar"
            Enabled         =   0   'False
            Height          =   330
            Index           =   0
            Left            =   60
            TabIndex        =   55
            ToolTipText     =   "Agregar Personal"
            Top             =   135
            Width           =   1065
         End
         Begin VB.CommandButton CmdPer 
            Caption         =   "&Seleccionar"
            Enabled         =   0   'False
            Height          =   330
            Index           =   1
            Left            =   1140
            TabIndex        =   54
            TabStop         =   0   'False
            ToolTipText     =   "Eliminar Personal"
            Top             =   135
            Width           =   1065
         End
         Begin VB.Label lblTotalGr 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   195
            Left            =   4710
            TabIndex        =   58
            Top             =   180
            Width           =   315
         End
      End
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   5830
         Picture         =   "FrmControlTarea1.frx":0B5B
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   49
         ToolTipText     =   "Cerrar"
         Top             =   50
         Width           =   195
      End
      Begin VB.Frame Frame7 
         Height          =   495
         Left            =   30
         TabIndex        =   50
         Top             =   270
         Width           =   6015
         Begin VB.OptionButton OptSeleccionar 
            Caption         =   "Seleccionar &Todos"
            Enabled         =   0   'False
            Height          =   225
            Index           =   2
            Left            =   60
            TabIndex        =   52
            Top             =   200
            Width           =   1785
         End
         Begin VB.OptionButton OptSeleccionar 
            Caption         =   "&Deseleccionar Todos"
            Enabled         =   0   'False
            Height          =   225
            Index           =   3
            Left            =   2010
            TabIndex        =   51
            Top             =   200
            Width           =   1965
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg4 
         Height          =   2100
         Left            =   90
         TabIndex        =   63
         Top             =   900
         Width           =   5925
         _cx             =   10451
         _cy             =   3704
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmControlTarea1.frx":0E47
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
         Caption         =   "Selección de Personal"
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
         TabIndex        =   62
         Top             =   45
         Width           =   1920
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   7
         X1              =   -60
         X2              =   6090
         Y1              =   3570
         Y2              =   3570
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   6
         X1              =   0
         X2              =   8295
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   4
         X1              =   6090
         X2              =   6090
         Y1              =   0
         Y2              =   3570
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   50
         Top             =   30
         Width           =   6000
      End
   End
   Begin VB.Frame Frm3 
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
      Height          =   3210
      Left            =   12030
      TabIndex        =   38
      Top             =   390
      Visible         =   0   'False
      Width           =   5460
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   5180
         Picture         =   "FrmControlTarea1.frx":0FBF
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   42
         ToolTipText     =   "Cerrar"
         Top             =   50
         Width           =   195
      End
      Begin VB.Frame Frame6 
         Height          =   435
         Left            =   90
         TabIndex        =   39
         Top             =   630
         Width           =   5280
         Begin VB.OptionButton OptSeleccionar 
            Caption         =   "Seleccionar Todos"
            Enabled         =   0   'False
            Height          =   225
            Index           =   0
            Left            =   60
            TabIndex        =   41
            Top             =   150
            Width           =   1785
         End
         Begin VB.OptionButton OptSeleccionar 
            Caption         =   "Deseleccionar Todos"
            Enabled         =   0   'False
            Height          =   225
            Index           =   1
            Left            =   2010
            TabIndex        =   40
            Top             =   150
            Width           =   1965
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg3 
         Height          =   1500
         Left            =   120
         TabIndex        =   43
         Top             =   1095
         Width           =   5190
         _cx             =   9155
         _cy             =   2646
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmControlTarea1.frx":12AB
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
         Height          =   555
         Left            =   120
         TabIndex        =   59
         Top             =   2550
         Width           =   5205
         Begin VB.CommandButton CmdTar 
            Caption         =   "&Mostrar Personal"
            Height          =   330
            Index           =   0
            Left            =   150
            TabIndex        =   60
            ToolTipText     =   "Agregar Personal"
            Top             =   150
            Width           =   1665
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Costo Total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   2730
            TabIndex        =   66
            Top             =   150
            Width           =   1050
         End
         Begin VB.Label LblCosto 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LblCosto"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4230
            TabIndex        =   65
            Top             =   150
            Width           =   615
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00800000&
            Height          =   285
            Index           =   5
            Left            =   2340
            Top             =   120
            Width           =   2580
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   195
            Left            =   4710
            TabIndex        =   61
            Top             =   180
            Width           =   315
         End
      End
      Begin VB.Label LblIdRec 
         AutoSize        =   -1  'True
         Caption         =   "LblIdRec"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4440
         TabIndex        =   47
         Top             =   390
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label LblProd 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblProd"
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   885
         TabIndex        =   46
         Top             =   330
         Width           =   4470
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         Height          =   195
         Left            =   60
         TabIndex        =   45
         Top             =   380
         Width           =   645
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   0
         X1              =   5430
         X2              =   5430
         Y1              =   -30
         Y2              =   3180
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   5
         X1              =   0
         X2              =   8295
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   4
         X1              =   -60
         X2              =   5430
         Y1              =   3180
         Y2              =   3180
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
         Left            =   105
         TabIndex        =   44
         Top             =   60
         Width           =   1770
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   50
         Top             =   30
         Width           =   5335
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7245
      Left            =   15
      TabIndex        =   6
      Top             =   360
      Width           =   11865
      _cx             =   20929
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
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6825
         Left            =   -12420
         TabIndex        =   7
         Top             =   375
         Width           =   11775
         Begin TrueOleDBGrid70.TDBGrid Dg3 
            Height          =   6465
            Left            =   45
            TabIndex        =   20
            Top             =   345
            Width           =   11745
            _ExtentX        =   20717
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
            Columns(1).Caption=   "Fch Trabajo"
            Columns(1).DataField=   "fchtra1"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Area"
            Columns(2).DataField=   "area"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Responsable"
            Columns(3).DataField=   "encargado"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Tipo"
            Columns(4).DataField=   "tipopago"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Estado"
            Columns(5).DataField=   "desestado"
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
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1535"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1455"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2117"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2037"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=4313"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=4233"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=7408"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=7329"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=3069"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2990"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14,.alignment=2"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
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
         Begin VB.Label lblperiodo 
            AutoSize        =   -1  'True
            Caption         =   "lblperiodo"
            BeginProperty Font 
               Name            =   "Book Antiqua"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   0
            Left            =   9705
            TabIndex        =   19
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Seguimiento de Tareas"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
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
            TabIndex        =   9
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label LblMes 
            Caption         =   "LblMes"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   300
            Left            =   8835
            TabIndex        =   8
            Top             =   30
            Width           =   1275
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Detalle de la Cuenta"
         Height          =   6825
         Left            =   45
         TabIndex        =   10
         Top             =   375
         Width           =   11775
         Begin VB.Frame FraProgreso 
            BorderStyle     =   0  'None
            Height          =   705
            Left            =   5040
            TabIndex        =   35
            Top             =   5490
            Visible         =   0   'False
            Width           =   5760
            Begin MSComctlLib.ProgressBar PgBar 
               Height          =   255
               Left            =   90
               TabIndex        =   36
               Top             =   360
               Width           =   5565
               _ExtentX        =   9816
               _ExtentY        =   450
               _Version        =   393216
               BorderStyle     =   1
               Appearance      =   0
               Scrolling       =   1
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00FFFFFF&
               BorderWidth     =   3
               Index           =   4
               X1              =   0
               X2              =   15
               Y1              =   15
               Y2              =   5070
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00000000&
               BorderWidth     =   2
               Index           =   3
               X1              =   -60
               X2              =   6360
               Y1              =   690
               Y2              =   690
            End
            Begin VB.Line Line11 
               BorderColor     =   &H00FFFFFF&
               BorderWidth     =   2
               X1              =   -150
               X2              =   5895
               Y1              =   0
               Y2              =   0
            End
            Begin VB.Line Line10 
               BorderColor     =   &H00000000&
               BorderWidth     =   2
               X1              =   5745
               X2              =   5745
               Y1              =   -90
               Y2              =   4800
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Importando"
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
               Left            =   90
               TabIndex        =   37
               Top             =   75
               Width           =   1020
            End
         End
         Begin VB.OptionButton OptTipoPago 
            Caption         =   "Destajo"
            Height          =   195
            Index           =   1
            Left            =   3660
            TabIndex        =   34
            Top             =   360
            Width           =   885
         End
         Begin VB.OptionButton OptTipoPago 
            Caption         =   "Horas"
            Height          =   195
            Index           =   0
            Left            =   2760
            TabIndex        =   33
            Top             =   360
            Width           =   795
         End
         Begin VB.CommandButton cb 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   6960
            Picture         =   "FrmControlTarea1.frx":136E
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   300
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Frame Frame4 
            Caption         =   "( Periodo )"
            Height          =   615
            Left            =   10020
            TabIndex        =   25
            Top             =   -15
            Width           =   1740
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
               Left            =   90
               TabIndex        =   26
               Top             =   315
               Width           =   1605
            End
         End
         Begin VB.CommandButton cb 
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   2040
            Picture         =   "FrmControlTarea1.frx":14A0
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   645
            Width           =   225
         End
         Begin VB.TextBox txt 
            BackColor       =   &H0080FF80&
            Height          =   315
            Index           =   0
            Left            =   3120
            TabIndex        =   17
            Tag             =   "null"
            Text            =   "txt(0)"
            Top             =   -45
            Visible         =   0   'False
            Width           =   1170
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
            Height          =   300
            Index           =   0
            Left            =   1080
            TabIndex        =   0
            Top             =   300
            Width           =   1230
            _ExtentX        =   2170
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
            Valor           =   "21/11/2007"
         End
         Begin VB.CommandButton cb 
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   6960
            Picture         =   "FrmControlTarea1.frx":15D2
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   645
            Width           =   225
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   1
            Left            =   6000
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   3
            Text            =   "txt_cb(1)"
            Top             =   615
            Width           =   1215
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   0
            Left            =   1080
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   2
            Text            =   "txt_cb(0)"
            Top             =   615
            Width           =   1215
         End
         Begin VB.TextBox txt_cb 
            Height          =   300
            Index           =   2
            Left            =   6000
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   1
            Text            =   "txt_cb(2)"
            Top             =   270
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   5370
            Left            =   45
            TabIndex        =   4
            Top             =   975
            Width           =   11730
            _cx             =   20690
            _cy             =   9472
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
            Rows            =   2
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmControlTarea1.frx":1704
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
            Height          =   525
            Left            =   0
            TabIndex        =   67
            Top             =   6270
            Width           =   11745
            Begin VB.CommandButton Cmd 
               Caption         =   "&Eliminar"
               Enabled         =   0   'False
               Height          =   300
               Index           =   1
               Left            =   1365
               TabIndex        =   74
               TabStop         =   0   'False
               ToolTipText     =   "Eliminar "
               Top             =   150
               Width           =   1275
            End
            Begin VB.CommandButton Cmd 
               Caption         =   "&Agregar"
               Enabled         =   0   'False
               Height          =   300
               Index           =   0
               Left            =   30
               TabIndex        =   73
               ToolTipText     =   "Agregar "
               Top             =   150
               Width           =   1275
            End
            Begin VB.CommandButton CmdUtil 
               Caption         =   "&Buscar"
               Height          =   300
               Index           =   0
               Left            =   2805
               TabIndex        =   72
               TabStop         =   0   'False
               ToolTipText     =   "Buscar Tareas Realizadas"
               Top             =   150
               Width           =   1275
            End
            Begin VB.CommandButton CmdUtil 
               Caption         =   "Exportar"
               Height          =   300
               Index           =   1
               Left            =   4140
               TabIndex        =   71
               TabStop         =   0   'False
               ToolTipText     =   "Exportar MSExcel"
               Top             =   150
               Width           =   1275
            End
            Begin VB.CommandButton Cmd 
               Caption         =   "Opciones"
               Enabled         =   0   'False
               Height          =   300
               Index           =   3
               Left            =   9765
               TabIndex        =   70
               TabStop         =   0   'False
               ToolTipText     =   "Opciones"
               Top             =   150
               Width           =   1395
            End
            Begin VB.CommandButton Cmd 
               Caption         =   "Buscar Tarea o Receta"
               Enabled         =   0   'False
               Height          =   300
               Index           =   2
               Left            =   7425
               TabIndex        =   69
               TabStop         =   0   'False
               ToolTipText     =   "Buscar Tarea o Receta"
               Top             =   150
               Width           =   2190
            End
            Begin VB.CommandButton Cmd 
               Caption         =   "Importar"
               Enabled         =   0   'False
               Height          =   300
               Index           =   4
               Left            =   5445
               TabIndex        =   68
               TabStop         =   0   'False
               ToolTipText     =   "Eliminar "
               Top             =   150
               Width           =   1275
            End
            Begin MSComDlg.CommonDialog Cmm 
               Left            =   6945
               Top             =   60
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
         End
         Begin VB.Label lblPesoTara 
            BackColor       =   &H00C0C0FF&
            Caption         =   "lblPesoTara(1)"
            Height          =   225
            Index           =   1
            Left            =   8790
            TabIndex        =   32
            Top             =   60
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label lblPesoTara 
            BackColor       =   &H00C0C0FF&
            Caption         =   "lblPesoTara(0)"
            Height          =   225
            Index           =   0
            Left            =   7650
            TabIndex        =   31
            Top             =   45
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label lbl_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cod(2)"
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
            Index           =   2
            Left            =   8220
            TabIndex        =   29
            Top             =   270
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "Contenedor"
            Height          =   195
            Index           =   2
            Left            =   4995
            TabIndex        =   28
            Top             =   405
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "Area"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   23
            Top             =   720
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
            Left            =   3285
            TabIndex        =   22
            Top             =   615
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Index           =   0
            Left            =   2565
            TabIndex        =   18
            Top             =   75
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbl_cod 
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cod(1)"
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
            Left            =   8475
            TabIndex        =   14
            Top             =   600
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Fch Trabajo"
            Height          =   195
            Index           =   8
            Left            =   60
            TabIndex        =   12
            Top             =   405
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle Seguimiento de Tarea"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   75
            TabIndex        =   11
            Top             =   15
            Width           =   11610
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "Responsable"
            Height          =   195
            Index           =   1
            Left            =   4995
            TabIndex        =   16
            Top             =   720
            Width           =   930
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
            Left            =   7215
            TabIndex        =   15
            Top             =   615
            Width           =   4545
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
            Left            =   2280
            TabIndex        =   24
            Top             =   615
            Width           =   2565
         End
         Begin VB.Label lbl_cb 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lbl_cb(2)"
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
            Left            =   7215
            TabIndex        =   30
            Top             =   270
            Visible         =   0   'False
            Width           =   2295
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   1005
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Cambiar Estado: PENDIENTE"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Cambiar Estado: PROCESADO"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Cambiar Estado: ANULADO"
               EndProperty
            EndProperty
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
         Left            =   6600
         Top             =   0
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
               Picture         =   "FrmControlTarea1.frx":1850
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlTarea1.frx":1D94
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlTarea1.frx":2126
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlTarea1.frx":22AA
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlTarea1.frx":26FE
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlTarea1.frx":2816
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlTarea1.frx":2D5A
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlTarea1.frx":329E
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlTarea1.frx":33B2
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlTarea1.frx":34C6
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlTarea1.frx":391A
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControlTarea1.frx":3A86
               Key             =   "IMG11"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "Copiar Celdas Seleccionadas"
      End
      Begin VB.Menu MenuEspc1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "Limpiar Celdas Seleccionadas"
      End
      Begin VB.Menu MenuEspc2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "Eliminar Filas Seleccionadas"
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "menu2"
      Visible         =   0   'False
      Begin VB.Menu menu2_1 
         Caption         =   "Copiar Celdas Seleccionadas"
      End
      Begin VB.Menu menu2_2 
         Caption         =   "Limpiar Celdas Seleccionadas"
      End
   End
End
Attribute VB_Name = "FrmControlTarea1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMCONTROLTAREA.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : PERMITE LAS ALTAS Y BAJAS DE LAS TAREAS REALIZADAS POR EL PERSONAL
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 29/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Public QueHace As Integer                     ' ESPECIFICA EN QUE ESTADO SE ENCUENTRA EL FORMULARIO
Dim Agregando As Boolean                      ' INDICA QUE SE ESTA AGREGANDO UNA FILA A LOS CONTROLES FLEXGRID
Dim SeEjecuto As Boolean                      ' VERIFICA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ
Dim RstFrm As New ADODB.Recordset             ' RECORDSET QUE ALMACENARA LOS DATOS DE LA TABLA pro_controltar
Dim mMesActivo  As Integer                    ' INDICA EL MES ACTIVO
Dim fOrdenLista As Boolean                    ' especfica el orden de la lista de la consulta
Dim mRowAdd As Double                         ' identificador unico por fila cuando se agrege una tarea
Dim mRowAddTara As Double                     ' identificador unico por fila cuando se agrege una tarea
Dim mIdRegistro&                              ' identificador del registro
Dim sPesoTara As Double
Public RstGrDet As New ADODB.Recordset        '
Public RstGrDetTara As New ADODB.Recordset    '
Private fActivarAutomaticoCantidad As Boolean ' permitira controlar la distribucion de las cantidades pro grupo
Private fActivarAutomaticoHora As Boolean     ' permitira controlar la distribucion de las horas pro grupo
Dim xHorIni As Date                           ' ESPECIFICA LA HORA DE INICIO DE UN PROCESO
Dim fCierrePeriodo As Boolean                 ' Indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)
Dim IdMenuActivo As Integer                   ' INDICA EL CODIGO DEL MENU ACTIVO

Dim RstTar As New ADODB.Recordset             ' Recordset especifico par el control de las tareas en linea
Dim RstPersonal As New ADODB.Recordset        ' Recordset especifico par el control de Personal

'Para mover el frame
Dim OrigFX As Long
Dim OrigFY As Long
Dim cSQL As String

Dim ESTADOPENDIENTE_ As Double
Dim ESTADOPROCESADO_ As Double
Dim ESTADOANULADO_ As Double

Private Sub chkOpcion_Click(Index As Integer)
    Select Case Index
        Case 2 ' tipo
            chkOpcion(3).Value = 0
        
        Case 3 '
            If chkOpcion(3).Value = 1 Then
                If chkOpcion(2).Value = 0 Then
                    chkOpcion(3).Value = 0
                End If
            End If
            
        Case 4 ' tarea
            chkOpcion(5).Value = 0
        
        Case 5 ' producto
            If chkOpcion(5).Value = 1 Then
                If chkOpcion(4).Value = 0 Then
                    chkOpcion(5).Value = 0
                End If
            End If
    End Select
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0 ' agregar
            pRegistroAdd
        
        Case 1 ' eliminar
            pRegistroDel
        
        Case 2 ' mostrar la consulta de tareas en las recetas
            pHabilitarBotonEditor 2, True
        
        Case 3 ' cuadro de opciones
            pHabilitarBotonEditor 3, True
        
        Case 4 ' importar
            pImportar
    End Select
End Sub

Private Sub cmdPer_click(Index As Integer)
    Select Case Index
        Case 0 ' Agregar
            Agregando = True
            Fg4_CellButtonClick Fg4.Row, 2
        Case 1 ' Seleccionar
            Agregando = False
            listarEmpleados
        Case 2 ' Eliminar
            Agregando = True
            eliminarRegistro
        Case 3 ' Eliminar todos
            Dim A As Integer
            Dim num As Integer
            
            num = Fg4.Rows - 1
            For A = 1 To num
                Agregando = False
                If Fg4.Rows > Fg4.FixedRows Then
                    Fg4.Select 1, 1
                    eliminarRegistro
                End If
            Next A
            pCargarDatos
        Case 4 ' Cargar Grupo
            agregarGrupo
    End Select
    Fg4.Select Fg4.Rows - 1, 2
End Sub

Private Sub agregarGrupo()
    Dim xCampos(3, 4) As String
    Dim nSQL As String
    Dim nSQLId As String
    Dim nTitulo As String
    Dim RstTmp As New ADODB.Recordset
    Dim RstAux As New ADODB.Recordset
    
    xCampos(0, 0) = "Nº Grupo":         xCampos(0, 1) = "nombre":     xCampos(0, 2) = "900":   xCampos(0, 3) = "C":
    xCampos(1, 0) = "Responsable":      xCampos(1, 1) = "encargado":  xCampos(1, 2) = "3500":  xCampos(1, 3) = "C":
    xCampos(2, 0) = "Nº Integrantes":   xCampos(2, 1) = "totper":     xCampos(2, 2) = "1400":  xCampos(2, 3) = "N":
                    
    nSQL = "SELECT pro_grupo.id as idgru, pro_grupo.num & '' as nombre , 'GRUPO Nº' &  pro_grupo.num as referencia, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS encargado, Count(pro_grupodet.idgrupo) AS totper " _
        + vbCr + " FROM (pro_grupo LEFT JOIN (pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) ON pro_grupo.idres = pro_emp.id) INNER JOIN pro_grupodet ON pro_grupo.id = pro_grupodet.idgrupo " _
        + vbCr + " GROUP BY pro_grupo.id, pro_grupo.num, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom];"
        
    nTitulo = "Buscando Grupos"
    
    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio, ""
    
    If RstTmp.State = 0 Then Exit Sub
    If RstTmp.RecordCount <> 0 Then
            
        If Fg4.Rows = Fg4.FixedRows Then Fg4.Rows = Fg4.Rows + 1

        ' generar la lista de personal para no considerar en la lista
        nSQLId = GRID_GENERAR_SQL_ID(Fg4, 7, "AND pla_empleados.id", "NOT IN")
        ' Se genera la consulta
        nSQL = "SELECT " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & " AS corr, pla_empleados.id AS idemp, pla_empleados.codigo, pro_grupo.id AS idgrupo, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, 0 AS cant, 0 AS cantbrut, -1 AS activo " _
            + vbCr + "FROM pro_grupo INNER JOIN ((pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_grupodet ON pro_emp.id = pro_grupodet.idper) ON pro_grupo.id = pro_grupodet.idgrupo " _
            + vbCr + "WHERE (((pla_empleados.fchcese) Is Null) AND ((pro_grupo.id)= " & RstTmp("idgru") & "))" & nSQLId
        
        RST_Busq RstAux, nSQL, xCon
        
        If RstAux.State = 0 Then Exit Sub
        
        If RstAux.RecordCount <> 0 Then
            While Not RstAux.EOF
                RstPersonal.AddNew
                RstPersonal("corr") = NulosN(RstAux("corr"))
                RstPersonal("idrec") = NulosN(lblIdRec.Caption)
                RstPersonal("idper") = NulosN(RstAux("idemp"))
                RstPersonal("codigo") = NulosC(RstAux("codigo"))
                RstPersonal("nombre") = NulosC(RstAux("nombres"))
                RstPersonal("activo") = RstAux("activo")
                RstPersonal("idunid") = 2
                RstPersonal("preuni") = NulosN(LblCosto.Caption)
                RstPersonal("imptot") = NulosN(LblCosto.Caption)
                
                RstAux.MoveNext
            Wend
            calcularCantidades
            RstPersonal.Filter = "corr = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & ""
            pCargarDatos False, True
        End If
    End If
    
    Set RstTmp = Nothing
    Set RstAux = Nothing
End Sub

Private Sub listarEmpleados()
    If QueHace = 3 Then Exit Sub
    
    Dim nSQL As String
    Dim nSQLId As String
    Dim nSQLTmp  As String
    Dim nTitulo As String
    Dim xform As New eps_librerias.FormSeleccion
    Dim xRs As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    'Dim num As Integer ' numero de registros que se van a agregar
        
    xCampos(0, 0) = "Cod. Empleado":        xCampos(0, 1) = "codemp":       xCampos(0, 2) = "2000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
    xCampos(1, 0) = "Apellidos y Nombres":  xCampos(1, 1) = "nombre":      xCampos(1, 2) = "5000":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
    xCampos(2, 0) = "Id":                   xCampos(2, 1) = "idemp":        xCampos(2, 2) = "1000":     xCampos(2, 3) = "C":    xCampos(2, 4) = "C"
    
    If Fg4.Rows = Fg4.FixedRows Then Fg4.Rows = Fg4.Rows + 1

    ' generar la lista de personal para no considerar en la lista
'''    RstPersonal.Filter = "corr = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & ""
'''    nSQLId = GENERAR_SQL_ID2(RstPersonal, "idper", " AND pla_empleados.id", "NOT IN", True)
    nSQLId = GRID_GENERAR_SQL_ID(Fg4, 7, "AND pla_empleados.id", "NOT IN")
    
    If Fg4.Row <= 0 Then Fg4.Row = 1
    ' generar semejanzas en la lista
    If NulosC(Fg4.TextMatrix(Fg4.Row, 2)) <> "" Then
        nSQLTmp = " AND UCASE([pla_empleados].[nombre]) LIKE '%" & UCase(NulosC(Fg4.TextMatrix(Fg4.Row, 2))) & "%'"
    End If
    
    ' generar la consulta
    nSQL = "SELECT 0 AS xsel, " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & " As corr, pla_empleados.codigo AS codemp, pla_empleados.id AS idemp, pla_empleados.nombre, -1 AS activo " _
        + vbCr + "FROM (pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper " _
        + vbCr + "Where (((pla_empleados.fchcese) Is Null) AND ((pro_empdet.idfun) = 6)) " & nSQLId & nSQLTmp _
        + vbCr + "ORDER BY pla_empleados.nombre;"
        
    nTitulo = "Buscando Personal"

    xform.SQLCad = nSQL
        
    xform.titulo = "Buscando Personal"
    Set xform.Coneccion = xCon
    Set xRs = Nothing
    Set xRs = xform.seleccionar(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            Dim A As Integer
            While Not xRs.EOF
                ' agregando los datos al rst temporal
                RstPersonal.AddNew
                
                RstPersonal("corr") = NulosN(xRs("corr"))
                RstPersonal("idrec") = NulosN(lblIdRec.Caption)
                RstPersonal("idper") = xRs("idemp")
                RstPersonal("codigo") = xRs("codemp")
                RstPersonal("nombre") = xRs("nombre")
                RstPersonal("activo") = xRs("activo")
                RstPersonal("idunid") = 2
                RstPersonal("preuni") = NulosN(LblCosto.Caption)
                RstPersonal("imptot") = NulosN(LblCosto.Caption)
                
                xRs.MoveNext
            Wend
            calcularCantidades
            RstPersonal.Filter = "corr = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & ""
            pCargarDatos False, True
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub calcularCantidades()
    Dim numper As Double
    Dim cantidad As Double
    
    If RstPersonal.State = 0 Then Exit Sub
    
    If chkAuto.Value = 1 Then
    
        RstPersonal.Filter = "corr = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & ""
        
        If RstPersonal.RecordCount <> 0 Then
            numper = 0
            cantidad = 0
            RstPersonal.MoveFirst
            ' Se halla el numero de personas activas
            While Not RstPersonal.EOF
                If RstPersonal("activo") = True Then
                    numper = numper + 1
                End If
                RstPersonal.MoveNext
            Wend
            
            cantidad = NulosN(Fg1.TextMatrix(Fg1.Row, 8))
            RstPersonal.MoveFirst
            While Not RstPersonal.EOF
                If RstPersonal("activo") = True Then
                
                    If IsDate(Fg1.TextMatrix(Fg1.Row, 6)) Then
                        RstPersonal("horini") = Format(Fg1.TextMatrix(Fg1.Row, 6), "HH:mm")
                    Else
                        RstPersonal("horini") = 0
                    End If
                    If IsDate(Fg1.TextMatrix(Fg1.Row, 7)) Then
                        RstPersonal("horfin") = Format(Fg1.TextMatrix(Fg1.Row, 7), "HH:mm")
                    Else
                        RstPersonal("horfin") = 0
                    End If
                    
                    RstPersonal("tothor") = 0
                    
                    If IsDate(Fg1.TextMatrix(Fg1.Row, 6)) And IsDate(Fg1.TextMatrix(Fg1.Row, 7)) Then
                        RstPersonal("difhor") = Format(CDate(RstPersonal("horfin")) - CDate(RstPersonal("horini")), "HH:mm")
                    Else
                        RstPersonal("difhor") = 0
                    End If
                    
                    RstPersonal("canpro") = cantidad / numper ' Se calcula la cantidad por persona
                Else
                    RstPersonal("horini") = 0
                    RstPersonal("horfin") = 0
                    RstPersonal("tothor") = 0
                    RstPersonal("difhor") = 0
                    RstPersonal("canpro") = 0
                End If
                
                RstPersonal.MoveNext
            Wend
        End If
        
    Else
        If Fg4.Rows = Fg4.FixedRows Then Exit Sub
        RstPersonal.Filter = "corr = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & " and idper = " & NulosN(Fg4.TextMatrix(Fg4.Row, 7))
        
        If RstPersonal.RecordCount <> 0 Then
            If IsDate(Fg4.TextMatrix(Fg4.Row, 3)) Then
                RstPersonal("horini") = Format(Fg4.TextMatrix(Fg4.Row, 3), "HH:mm")
            Else
                RstPersonal("horini") = 0
            End If
            If IsDate(Fg4.TextMatrix(Fg4.Row, 4)) Then
                RstPersonal("horfin") = Format(Fg4.TextMatrix(Fg4.Row, 4), "HH:mm")
            Else
                RstPersonal("horfin") = 0
            End If
            
            RstPersonal("tothor") = 0
            
            If IsDate(Fg4.TextMatrix(Fg4.Row, 3)) And IsDate(Fg4.TextMatrix(Fg4.Row, 4)) Then
                RstPersonal("difhor") = Format(CDate(RstPersonal("horfin")) - CDate(RstPersonal("horini")), "HH:mm")
            Else
                RstPersonal("difhor") = 0
            End If
            
            RstPersonal("canpro") = NulosN(Fg4.TextMatrix(Fg4.Row, 5))
        End If
        
        '--aplicando el filtro para mostrar personal de la linea seleccionada
        RstPersonal.Filter = "corr = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & ""
        Agregando = True
        
        '--Calcular Acumular cantidad en registro principal solo cuando chek esta desactivado
        Fg1.TextMatrix(Fg1.Row, 8) = RstRegistroSumar(RstPersonal, "canpro", "activo", "-1")
        Agregando = False
    End If

''    pCargarDatos False, True
    
End Sub

Private Sub eliminarRegistro()
    If Fg4.Row < 1 Then
        MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg4.SetFocus
        Exit Sub
    End If
    
    If Fg4.Rows = 1 Then
        MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg4.SetFocus
        Exit Sub
    End If
    
    If Agregando Then
        If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    End If
    
    If RstPersonal.RecordCount <> 0 Then RstPersonal.MoveFirst
    
    Do While Not RstPersonal.EOF
        If RstPersonal.RecordCount = 0 Then Exit Do
        If NulosN(RstPersonal("idper")) = NulosN(Fg4.TextMatrix(Fg4.Row, 7)) Then
            RstPersonal.Delete
            Exit Do
        End If
        RstPersonal.MoveNext
    Loop
    calcularCantidades
    pCargarDatos False, True
End Sub

Private Sub cmdTar_click(Index As Integer)
    If Index = 0 Then
        If RstPersonal.State <> 0 Then
            RstPersonal.Filter = adFilterNone
            If RstPersonal.RecordCount <> 0 Then
                RstPersonal.Filter = "corr = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & ""
            End If
        Else
            Dim cSQL As String
            Dim RstTmp As New ADODB.Recordset
        
            cSQL = "SELECT " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & " As corr, pro_controltardetgr.idrec, pro_controltardetgr.idper, pla_empleados.codigo, pla_empleados.nombre, pro_controltardetgr.activo, pro_controltardetgr.horini, pro_controltardetgr.horfin, pro_controltardetgr.tothor, pro_controltardetgr.difhor, pro_controltardetgr.canpro, pro_controltardetgr.idunid, pro_controltardetgr.preuni, pro_controltardetgr.imptot " _
                    + vbCr + "FROM pro_controltardetgr LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id;"
                    
            RST_Busq RstTmp, cSQL, xCon
            
            DEFINIR_RST_TMP RstPersonal, RstTmp
        End If
        
        If RstPersonal.RecordCount <> 0 Then
            pCargarDatos True, False ' Se refrescan los costos
            modificarRst RstPersonal, False, True ' se modifica el personal con los nuevos costos
        End If
        
        Frm4.Visible = True
        
    End If
End Sub

Private Sub CmdTarea_Click(Index As Integer)
    pHabilitarBotonEditor 2, False
End Sub

Private Sub CmdUtil_Click(Index As Integer)
    Select Case Index
        Case 0 ' buscar registros
            pBuscarVSFlexGrid
        Case 1 ' exportar msexcel vsflexgrid
            pExportarVSFlexGrid
    End Select
End Sub

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace <> 3 Then
'            If Index = 0 Then PopupMenu Menu3
'            If Index = 1 Then PopupMenu menu2
        End If
    End If
End Sub

Private Sub Dg3_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg3_FilterChange()
    TDB_FiltroGenerar Dg3, RstFrm
End Sub

Private Sub Dg3_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA EN FORMA ASCENDENTE O DESCENDENTE LAS COLUMNAS DEL CONTROL Dg3
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstFrm.Sort = CStr(Dg3.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstFrm("id")), xCon
    End If
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    If Agregando = True Then Exit Sub
    If Row = 0 Then Exit Sub
   
    fActivarAutomaticoHora = False
    fActivarAutomaticoCantidad = False
    
    Select Case Col
        Case 2 ' tipo persona
            ' actualizando el codigo
            If NulosN(Fg1.TextMatrix(Row, 11)) <> NulosN(Fg1.TextMatrix(Row, 2)) And NulosN(Fg1.TextMatrix(Row, 11)) <> 0 Then
                ' eliminar los registros del grupo anterior
                If NulosN(Fg1.TextMatrix(Row, 12)) <> 0 Then RstRegistroEliminar RstGrDet, "codigo", NulosN(Fg1.TextMatrix(Row, 16)), True
                Fg1.TextMatrix(Row, 3) = "": ' personal / grupo / linea
                Fg1.TextMatrix(Row, 12) = "" ' idref personal / grupo / linea
            End If
            Fg1.TextMatrix(Row, 11) = NulosN(Fg1.TextMatrix(Row, 2))
            
        Case 3 ' personal(individual/grupal)
            ' limpiar CodRef
            If NulosC(Fg1.TextMatrix(Row, Col)) = "" Then Fg1.TextMatrix(Row, 12) = ""
        
        Case 4 ' tarea
            ' limpiar CodTarea
            If NulosC(Fg1.TextMatrix(Row, Col)) = "" Then
                Fg1.TextMatrix(Row, 14) = "" ' idtarea
                Fg1.TextMatrix(Row, 5) = ""  ' producto
                Fg1.TextMatrix(Row, 13) = "" ' idrec
            End If
            
        Case 5 ' producto
            ' limpiar CodReceta
            If NulosC(Fg1.TextMatrix(Row, Col)) = "" Then Fg1.TextMatrix(Row, 13) = ""
            
        Case 6, 7
            If Fg1.TextMatrix(Row, Col) = "" Then
                GoTo Continuar1
            End If
            If Fg1.TextMatrix(Row, Col) = "  :  " Then
                Fg1.TextMatrix(Row, Col) = "":  GoTo Continuar1:
            End If
            If IsDate(Fg1.TextMatrix(Row, Col)) = False Then
                MsgBox "El valor ingresado no es una Hora correcta", vbCritical, xTitulo
                Fg1.TextMatrix(Row, Col) = ""
            Else
                Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), FORMAT_HORA_SIN_SEGUNDO)
                pConvertHora Row
Continuar1:
                ' actualizar si es grupo
                If NulosN(Fg1.TextMatrix(Row, 2)) = 2 Then
                    If BuscarFrm("FrmControlTareaGr1", True) = True Then
                        If FrmControlTareaGr1.fDesactivarAuto = False Then fActivarAutomaticoHora = True
                    Else
                        fActivarAutomaticoHora = True
                    End If
                    Fg1_RowColChange
                    fActivarAutomaticoHora = False
                End If
                
                '***********************************************************************************
                If NulosN(Fg1.TextMatrix(Row, 2)) = 3 And chkAuto.Value = 1 Then
                    calcularCantidades
                    pCargarDatos False, True
                End If
                '***********************************************************************************
            End If

        Case 8 ' cantidad
            If NulosC(Fg1.TextMatrix(Row, Col)) = "" Then
                ' actualizar si es grupo
                fActivarAutomaticoCantidad = True
                If NulosN(Fg1.TextMatrix(Row, 2)) = 2 Then Fg1_RowColChange
                fActivarAutomaticoCantidad = False
            Else
                If IsNumeric(Fg1.TextMatrix(Row, Col)) = False Then
                    MsgBox "El valor ingresado no es numérico", vbCritical, xTitulo
                    Fg1.TextMatrix(Row, Col) = ""
                Else
                    Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), FORMAT_CANTIDAD)
                    
                    '***********************************************************************************
                    If NulosN(Fg1.TextMatrix(Row, 2)) = 3 And chkAuto.Value = 1 Then
                        calcularCantidades
                        pCargarDatos False, True
                    End If
                    '***********************************************************************************
                    
                    ' actualizar si es grupo
                    If NulosN(Fg1.TextMatrix(Row, 2)) = 2 Then
                        If BuscarFrm("FrmControlTareaGr1", True) = True Then
                            If FrmControlTareaGr1.fDesactivarAuto = False Then
                                If NulosN(Fg1.TextMatrix(Row, Col)) <> NulosN(Fg1.TextMatrix(Row, 17)) Then fActivarAutomaticoCantidad = True
                            End If
                        Else
                            If NulosN(Fg1.TextMatrix(Row, Col)) <> NulosN(Fg1.TextMatrix(Row, 17)) Then fActivarAutomaticoCantidad = True
                        End If
                        Fg1.TextMatrix(Row, 17) = Fg1.TextMatrix(Row, Col)
                        Fg1_RowColChange
                    End If
                End If
            End If
            
        Case 9 ' unidad de medida
            ' limpiar CodTarea
            If NulosC(Fg1.TextMatrix(Row, Col)) = "" Then Fg1.TextMatrix(Row, 15) = ""
            pConvertHora Row

    End Select
    Exit Sub

error:
    SHOW_ERROR Me.Name, "Fg1_CellChanged ( " & Row & "," & Col & ")"
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col <> 1 And Col <> 3 And Col <> 4 And Col <> 5 And Col <> 8 And Col <> 9 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLTmp As String
    Dim nSQLNotId As String
    Dim nTitulo As String
    
    On Error GoTo error
        
    Select Case Col
        Case 1 ' Num. Reg. Prod.
            ReDim xCampos(6, 4) As String
            
            'descripcion                        'campo                              'tamaño                         'tipo = Numerico, caracter, fecha
            xCampos(0, 0) = "Num. Prod.":       xCampos(0, 1) = "numparte":          xCampos(0, 2) = "1200":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Descripcion":      xCampos(1, 1) = "despro":       xCampos(1, 2) = "3500":         xCampos(1, 3) = "C"
            xCampos(2, 0) = "Fech. Pro.":       xCampos(2, 1) = "dia":          xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
            xCampos(3, 0) = "Hor. Pro.":        xCampos(3, 1) = "horini":       xCampos(3, 2) = "900":          xCampos(3, 3) = "C"
            xCampos(4, 0) = "U.M":              xCampos(4, 1) = "abrev":        xCampos(4, 2) = "500":          xCampos(4, 3) = "C"
            xCampos(5, 0) = "Cantidad":         xCampos(5, 1) = "cantidad":     xCampos(5, 2) = "1000":         xCampos(5, 3) = "N"
                
            nSQL = "SELECT pro_produccion.dia, pro_receta.iditem, alm_inventario.descripcion AS despro, pro_producciondet.idrec, pro_receta.codrec, pro_producciondet.horini, pro_producciondet.horfin, pro_producciondet.numparte, pro_producciondet.cantidad, pro_receta.idunimed, mae_unidades.abrev, pro_emp.idemp AS idresp, pla_empleados.nombre, pro_producciondet.corr AS idregprod " _
                    + vbCr + "FROM pro_produccion LEFT JOIN (((((pro_producciondet LEFT JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_receta.idunimed = mae_unidades.id) LEFT JOIN pro_emp ON pro_producciondet.idres = pro_emp.id) LEFT JOIN pla_empleados ON pro_emp.idemp = pla_empleados.id) ON pro_produccion.id = pro_producciondet.idpro;"

            nTitulo = "Buscando Reg. Prod."
            
        Case 2 ' tipo persona
            
        Case 3 ' personal / nº grupo
        
            If NulosN(Fg1.TextMatrix(Row, 2)) = 0 Then ' indivual
                MsgBox "Seleccione el Tipo Individual o Grupal o Lineal", vbExclamation, xTitulo
                Fg1.Col = 2
                Fg1.SetFocus
                Exit Sub
            ElseIf NulosN(Fg1.TextMatrix(Row, 2)) = 1 Then ' individual
                If NulosC(Fg1.TextMatrix(Row, Col)) <> "" Then
                    nSQLTmp = " AND UCASE([pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom]) LIKE '%" & UCase(NulosC(Fg1.TextMatrix(Row, Col))) & "%'"
                End If
            
                ReDim xCampos(4, 4) As String
                xCampos(0, 0) = "Cod. Empleado":        xCampos(0, 1) = "codemp":      xCampos(0, 2) = "2000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
                xCampos(1, 0) = "Apellidos y Nombres":  xCampos(1, 1) = "nombre":      xCampos(1, 2) = "4500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
                xCampos(2, 0) = "Fch. Nac":             xCampos(2, 1) = "fchnac":      xCampos(2, 2) = "1000":     xCampos(2, 3) = "C":    xCampos(2, 4) = "C"
                xCampos(3, 0) = "Id":                   xCampos(3, 1) = "idemp":       xCampos(3, 2) = "600":      xCampos(3, 3) = "N":    xCampos(3, 4) = "C"
                
                nSQL = "SELECT pla_empleados.id AS idemp, pla_empleados.nombre, pla_empleados.fchnac, pla_empleados.codigo AS codemp " _
                    + vbCr + " FROM (pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper " _
                    + vbCr + " WHERE pla_empleados.fchcese is null and (((pro_empdet.idfun) = 6 )) " & nSQLTmp _
                    + vbCr + " ORDER BY pla_empleados.nombre ; "

                nTitulo = "Buscando Personal"
            
            ElseIf NulosN(Fg1.TextMatrix(Row, 2)) = 2 Then ' grupal
                ReDim xCampos(3, 4) As String
                xCampos(0, 0) = "Nº Grupo":         xCampos(0, 1) = "nombre":     xCampos(0, 2) = "900":   xCampos(0, 3) = "C":
                xCampos(1, 0) = "Responsable":      xCampos(1, 1) = "encargado":  xCampos(1, 2) = "3500":  xCampos(1, 3) = "C":
                xCampos(2, 0) = "Nº Integrantes":   xCampos(2, 1) = "totper":     xCampos(2, 2) = "1400":  xCampos(2, 3) = "N":
                        
                If NulosN(Fg1.TextMatrix(Row, 12)) <> 0 Then nSQLNotId = " WHERE pro_grupo.id <> " & NulosN(Fg1.TextMatrix(Row, 12))
                        
                nSQL = "SELECT pro_grupo.id as idgru, pro_grupo.num & '' as nombre , 'GRUPO Nº' &  pro_grupo.num as referencia, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS encargado, Count(pro_grupodet.idgrupo) AS totper " _
                    + vbCr + " FROM (pro_grupo LEFT JOIN (pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) ON pro_grupo.idres = pro_emp.id) INNER JOIN pro_grupodet ON pro_grupo.id = pro_grupodet.idgrupo " _
                    + vbCr + nSQLNotId & " GROUP BY pro_grupo.id, pro_grupo.num, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom];"
                nTitulo = "Buscando Grupos"
            ElseIf NulosN(Fg1.TextMatrix(Row, 2)) = 3 Then ' lineal
                '--pendiente
                Exit Sub
            End If
        
        Case 4 ' de la tarea
            '--Si es lineal no se selecciona la tarea, esta dependera del producto seleccionado
            If NulosN(Fg1.TextMatrix(Row, 2)) = 3 Then ' lineal
                Exit Sub
            End If
        
            ReDim xCampos(4, 4) As String
            xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":     xCampos(0, 2) = "4500":    xCampos(0, 3) = "C"
            xCampos(1, 0) = "Abrev":        xCampos(1, 1) = "nomcorto":   xCampos(1, 2) = "2300":    xCampos(1, 3) = "C"
            xCampos(2, 0) = "Diverso":      xCampos(2, 1) = "diverso":    xCampos(2, 2) = "700":     xCampos(2, 3) = "C"
            xCampos(3, 0) = "Id":           xCampos(3, 1) = "id":         xCampos(3, 2) = "600":     xCampos(3, 3) = "N"
            
            If NulosC(Fg1.TextMatrix(Row, Col)) <> "" Then
                nSQLTmp = " AND (UCASE(pro_tareas.descripcion) LIKE '%" & UCase(NulosC(Fg1.TextMatrix(Row, Col))) & "%' OR UCASE(pro_tareas.abrev) LIKE '%" & UCase(NulosC(Fg1.TextMatrix(Row, Col))) & "%' ) "
            End If
            
            ' si hay area seleccionada o que el filtro de areas seleccioanadas este desacivadas
            If NulosN(lbl_cod(0).Caption) <> 0 And chkOpcion(0).Value = 0 Then
                nSQL = "SELECT pro_tareas.id, pro_tareas.codigo, pro_tareas.descripcion AS nombre,pro_tareas.abrev AS nomcorto, mae_unidades.id AS idunimed, mae_unidades.abrev, IIf([pro_tareas].[diverso]=-1,'Si','No') AS diverso " _
                    + vbCr + " FROM (mae_unidades RIGHT JOIN (pro_tareas LEFT JOIN pro_areadet ON pro_tareas.id = pro_areadet.idtar) ON mae_unidades.id = pro_tareas.idunimed) LEFT JOIN pro_area ON pro_areadet.idar = pro_area.id " _
                    + vbCr + " WHERE pro_areadet.activo = -1 And pro_area.idarea = " & NulosN(lbl_cod(0).Caption)
            Else ' no hay area seleccionada
                nSQLTmp = Replace(nSQLTmp, "AND", "WHERE")
                nSQL = "SELECT pro_tareas.id, pro_tareas.codigo, pro_tareas.descripcion AS nombre, pro_tareas.abrev AS nomcorto, mae_unidades.id AS idunimed, mae_unidades.abrev, IIf([pro_tareas].[diverso]=-1,'Si','No') AS diverso " _
                    + vbCr + " FROM mae_unidades RIGHT JOIN pro_tareas ON mae_unidades.id = pro_tareas.idunimed  "
            End If
            nSQL = nSQL & nSQLTmp
            
            nTitulo = "Buscando Tareas"
        
        Case 5 ' del producto
            '--validar el ingreso de tarea para seleccionar el producto
            '--si es lineal no es necesario seleccionar la tarea, luego de seleccionar el producto se seleccionara de una lista las tareas
            If NulosN(Fg1.TextMatrix(Row, 2)) <> 3 Then
                If NulosN(Fg1.TextMatrix(Row, 14)) = 0 Then '--id tarea
                    MsgBox "Falta Especificar la Tarea", vbExclamation, xTitulo
                    Fg1.Col = 4
                    Fg1.SetFocus
                    Exit Sub
                End If
            End If
            
            ReDim xCampos(3, 4) As String
            xCampos(0, 0) = "Código":       xCampos(0, 1) = "codpro":   xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
            xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "nombre":   xCampos(1, 2) = "5000":    xCampos(1, 3) = "C"
            xCampos(2, 0) = "CodReceta":    xCampos(2, 1) = "codrec":   xCampos(2, 2) = "1200":    xCampos(2, 3) = "C"
            
            If NulosC(Fg1.TextMatrix(Row, Col)) <> "" Then
                nSQLTmp = " AND UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(Fg1.TextMatrix(Row, Col))) & "%'"
            End If
            
            '--mostrar listado segun filtro de tarea
            If NulosN(Fg1.TextMatrix(Row, 2)) <> 3 Then
                nSQL = "SELECT DISTINCT alm_inventario.codpro, alm_inventario.descripcion as nombre, pro_receta.codrec, pro_receta.iditem, pro_receta.id AS idrec " _
                    + vbCr + " FROM alm_inventario INNER JOIN pro_receta ON alm_inventario.id = pro_receta.iditem " _
                    + vbCr + " WHERE pro_receta.id IN (SELECT pro_recetatar.idrec FROM pro_recetatar WHERE pro_recetatar.idtar= " & NulosN(Fg1.TextMatrix(Row, 14)) & ") " & nSQLTmp _
                    + vbCr + " ORDER BY alm_inventario.descripcion; "
            Else
            '--mostrar listado de producto solo cuando tipo de trabajo sea lineal

                nSQL = "SELECT DISTINCT alm_inventario.codpro, alm_inventario.descripcion AS nombre, pro_receta.codrec, pro_receta.iditem, pro_receta.id AS idrec, pro_receta.idunimed, mae_unidades.abrev " _
                    + vbCr + " FROM (alm_inventario RIGHT JOIN pro_receta ON alm_inventario.id = pro_receta.iditem) LEFT JOIN mae_unidades ON pro_receta.idunimed = mae_unidades.id " _
                    + vbCr + " WHERE pro_receta.id IN (SELECT pro_recetatar.idrec FROM pro_recetatar) " & nSQLTmp _
                    + vbCr + " ORDER BY alm_inventario.descripcion; "
            End If
    
            nTitulo = "Buscando Productos"
    
        Case 9 ' unidad de medida
            ReDim xCampos(2, 4) As String
            xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":   xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Abrev":        xCampos(1, 1) = "abrev":    xCampos(1, 2) = "800":    xCampos(1, 3) = "C"
            xCampos(2, 0) = "Id":           xCampos(2, 1) = "id":       xCampos(2, 2) = "600":    xCampos(2, 3) = "N"
        
            nSQL = "SELECT mae_unidades.id, mae_unidades.descripcion as nombre, mae_unidades.abrev FROM mae_unidades;"
        
        Case 8
            pHabilitarBotonEditor 1, True
            Exit Sub
        
        Case Else
            Exit Sub
    End Select
    
    If Col = 1 Then
        CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), nTitulo, "dia", "numparte", CualquierParte
    Else
        CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio, ""
    End If

    If RstTmp.State = 0 Then GoTo SALIR
    If RstTmp.EOF = True Or RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR
    Agregando = True
    
    If Col = 1 Then ' Num. Reg. Prod.
        Fg1.TextMatrix(Row, 1) = NulosC(RstTmp("numparte"))
        Fg1.TextMatrix(Row, 2) = 3
        Fg1.TextMatrix(Row, 11) = NulosN(Fg1.TextMatrix(Row, 2))
        
        Fg1.TextMatrix(Row, 5) = NulosC(RstTmp("despro"))
        Fg1.TextMatrix(Row, 13) = NulosN(RstTmp("idrec"))
                       
        ' agregando la unidad por defecto
        Fg1.TextMatrix(Row, 9) = NulosC(RstTmp("abrev"))
        Fg1.TextMatrix(Row, 15) = NulosN(RstTmp("idunimed"))
        ' Se define o se llena las tareas para esa receta segun su unidad de medida
        definir_llenar_Tareas NulosN(RstTmp("idrec")), NulosN(RstTmp("idunimed"))
        
        Fg1.TextMatrix(Row, 22) = NulosN(RstTmp("idregprod"))
        
    ElseIf Col = 2 Then ' tipo persona
        'Fg1.TextMatrix(Row, Col) = NulosC(RstTmp.Fields("nombre"))
    ElseIf Col = 3 Then ' persona / grupo
        If NulosN(Fg1.TextMatrix(Row, 2)) = 2 Then '--solo grupos
            ' si el grupo inicial es diferente al grupo actual
            If NulosN(Fg1.TextMatrix(Row, 12)) <> NulosN(RstTmp.Fields("idgru")) Then
                ' eliminar los registros del grupo anterior
                RstGrDet.Filter = ""
                If NulosN(Fg1.TextMatrix(Row, 12)) <> 0 Then RstRegistroEliminar RstGrDet, "codigo", NulosN(Fg1.TextMatrix(Row, 16)), True
                ' cargar los datos del grupo
                pCargarDatosRstTemp 0, NulosN(RstTmp.Fields("idgru")), NulosN(Fg1.TextMatrix(Row, 16)), False
                Fg1.TextMatrix(Row, 12) = NulosN(RstTmp.Fields("idgru"))
                ' mostrar los datos en la ventana
                Agregando = False
                fActivarAutomaticoCantidad = True
                Fg1_RowColChange
                fActivarAutomaticoCantidad = False
                Agregando = True
            End If
            Fg1.TextMatrix(Row, Col) = NulosC(RstTmp.Fields("referencia"))
        Else
            Fg1.TextMatrix(Row, Col) = NulosC(RstTmp.Fields("nombre"))
            Fg1.TextMatrix(Row, 12) = NulosN(RstTmp.Fields("idemp"))
        End If
        Fg1.Col = 4
    ElseIf Col = 4 Then ' tarea
        ' si la tarea es diferente => limpiar producto
        If NulosN(Fg1.TextMatrix(Row, 13)) <> 0 And (NulosN(Fg1.TextMatrix(Row, 14)) <> NulosN(RstTmp.Fields("id"))) Then
            Fg1.TextMatrix(Row, 5) = ""
            Fg1.TextMatrix(Row, 13) = ""
        End If
    
        Fg1.TextMatrix(Row, Col) = NulosC(RstTmp.Fields("nomcorto"))
        Fg1.TextMatrix(Row, 14) = NulosN(RstTmp.Fields("id"))
        ' agregando la unidad por defecto
        Fg1.TextMatrix(Row, 9) = NulosC(RstTmp.Fields("abrev"))
        Fg1.TextMatrix(Row, 15) = NulosN(RstTmp.Fields("idunimed"))
        pConvertHora Row
        Fg1.Col = 5
    ElseIf Col = 5 Then ' producto
    
        Fg1.TextMatrix(Row, Col) = NulosC(RstTmp.Fields("nombre"))
        Fg1.TextMatrix(Row, 13) = NulosN(RstTmp.Fields("idrec"))
        
        '****************************************************************************************************************
        ' Si es lineal se muestran las tareas de esa linea
        If NulosN(Fg1.TextMatrix(Row, 2)) = 3 Then
            Dim RstTemp As New ADODB.Recordset
            
            ' agregando la unidad por defecto
            Fg1.TextMatrix(Row, 9) = NulosC(RstTmp.Fields("abrev"))
            Fg1.TextMatrix(Row, 15) = NulosN(RstTmp.Fields("idunimed"))
            ' Se define o se llena las tareas para esa receta segun su unidad de medida
            definir_llenar_Tareas NulosN(RstTmp.Fields("idrec")), NulosN(RstTmp.Fields("idunimed"))
        End If
        '****************************************************************************************************************
        Fg1.Col = 6
    ElseIf Col = 9 Then ' unidad de medida
        Fg1.TextMatrix(Row, Col) = NulosC(RstTmp.Fields("abrev"))
        Fg1.TextMatrix(Row, 15) = NulosN(RstTmp.Fields("id"))
        pConvertHora Row
        Fg1.Col = 10
        
        '********************************************************************
        ' Se define o se llena las tareas para esa receta segun su unidad de medida
        definir_llenar_Tareas NulosN(Fg1.TextMatrix(Row, 13)), NulosN(RstTmp.Fields("id"))
        '********************************************************************
        
    End If
    Agregando = False
    Set RstTmp = Nothing
    Exit Sub

SALIR:
    If Col = 3 Then     ' persona / grupo
        Fg1.Col = 3
    ElseIf Col = 4 Then ' producto
        Fg1.Col = 4
    ElseIf Col = 5 Then ' tarea
        Fg1.Col = 5
    ElseIf Col = 9 Then ' unidad de medida
        Fg1.Col = 9
    End If
    Fg1.TextMatrix(Row, Col) = ""
    Fg1.SetFocus
    Set RstTmp = Nothing
    Agregando = False
    Exit Sub

error:
    Set RstTmp = Nothing
    Agregando = False
    SHOW_ERROR Me.Name, "Fg1_CellButtonClick(" & Row & "," & Col & ")"
End Sub

Private Sub definir_llenar_Tareas(IDRECETA_ As Double, Optional IDUNIDAD_ As Double = 2, _
                                    Optional CORR_ As Double = 0, Optional PRODUCTO_ As String = "")
                                    
    Dim nSQL As String
    Dim RstTemp As New ADODB.Recordset
    
    If CORR_ = 0 Then
        CORR_ = NulosN(Fg1.TextMatrix(Fg1.Row, 16))
    End If
    
    If PRODUCTO_ = "" Then
        PRODUCTO_ = NulosC(Fg1.TextMatrix(Fg1.Row, 4))
    End If
    
    nSQL = "SELECT " & CORR_ & " AS corr, -1 AS activo, pro_receta.id AS idrec, pro_receta.iditem AS idpro, pro_recetatar.orden, alm_inventario.descripcion AS nompro, pro_tareas.codigo AS codtar, pro_recetatar.idtar, pro_tareas.descripcion AS destar, costo.costo " _
        + vbCr + "FROM ((pro_receta INNER JOIN (pro_recetatar INNER JOIN pro_tareas ON pro_recetatar.idtar = pro_tareas.id) ON pro_receta.id = pro_recetatar.idrec) INNER JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) LEFT JOIN " _
        + vbCr + "( " _
        + vbCr + "SELECT pro_costo.idref, pro_costodet.idtar, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.costo " _
        + vbCr + "FROM pro_tareas INNER JOIN (pro_costo INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
        + vbCr + "Where (((pro_costo.idref) = " & NulosN(IDRECETA_) & ") And ((pro_costodet.idunimed) = " & NulosN(IDUNIDAD_) & ") And ((pro_costo.Tipo) = 1)) " _
        + vbCr + ") " _
        + vbCr + "AS costo ON (pro_recetatar.idtar = costo.idtar) AND (pro_recetatar.idrec = costo.idref) " _
        + vbCr + "WHERE (((pro_receta.id)= " & NulosN(IDRECETA_) & "));"
    
    RST_Busq RstTemp, nSQL, xCon
    
    If RstTemp.State = 0 Then Exit Sub
    
    If RstTar.State = 0 Then
        DEFINIR_RST_TMP RstTar, RstTemp
    Else
        RstTar.Filter = adFilterNone ' Se quitan los filtros
        RstTar.Filter = "corr = " & CORR_ & "" ' Se filtra el correlativo
        
        If RstTar.RecordCount <> 0 Then
            ' Se eliminan los registros anteriores
            RstRegistroEliminar RstTar, "corr", CStr(CORR_), True
        End If
    End If
    
    If RstTemp.RecordCount <> 0 Then
        LblProd.Caption = PRODUCTO_
        lblIdRec.Caption = NulosN(IDRECETA_)
        CARGAR_RST_TMP RstTar, RstTemp
    Else
        LblProd.Caption = ""
        lblIdRec.Caption = ""
        LblCosto.Caption = ""
    End If
    RstTar.Filter = "corr = " & CORR_ & "" ' Se filtra el correlativo
    pCargarDatos True, False
End Sub

Private Sub modificarRst(ByRef Rst As ADODB.Recordset, Optional tareas As Boolean = True _
                        , Optional personal As Boolean = False)
    Dim A As Integer
    With Rst
        If tareas Then
            For A = 1 To Fg3.Rows - 1
                ' Se filtran los datos correspondientes
                .Filter = "idtar = " & NulosN(Fg3.TextMatrix(A, 6)) & " And corr = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & ""
                If .RecordCount <> 0 Then
                    Fg3.Select A, 2
                    If Fg3.CellChecked = flexChecked Then .Fields("activo") = True Else .Fields("activo") = False
                End If
            Next A
            
            .Filter = adFilterNone
            .Filter = "corr = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & ""
        End If
        
        If personal Then
'            For A = 1 To Fg4.Rows - 1
'                ' Se filtran los datos correspondientes
'                .Filter = "idper = " & NulosN(Fg4.TextMatrix(A, 7)) & " And corr = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & ""
'                If .RecordCount <> 0 Then
'                    Fg4.Select A, 2
'                    If Fg4.CellChecked = flexChecked Then .Fields("activo") = True Else .Fields("activo") = False
'                    '--asignar costo de tareas seleccionadas
'                    .Fields("preuni") = LblCosto.Caption
'                End If
'            Next A
            
            'For A = 1 To Fg4.Rows - 1
                ' Se filtran los datos correspondientes
                .Filter = "idper = " & NulosN(Fg4.TextMatrix(Fg4.Row, 7)) & " And corr = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & ""
                If .RecordCount <> 0 Then
                    Agregando = True
                    Fg4.Select Fg4.Row, 2
                    If Fg4.CellChecked = flexChecked Then
                        .Fields("activo") = True
                    Else
                        .Fields("activo") = False
                        .Fields("horini") = 0
                        .Fields("horfin") = 0
                        .Fields("canpro") = 0
                        Fg4.TextMatrix(Fg4.Row, 3) = ""
                        Fg4.TextMatrix(Fg4.Row, 4) = ""
                        Fg4.TextMatrix(Fg4.Row, 5) = ""
                    End If
                    '--asignar costo de tareas seleccionadas
                    .Fields("preuni") = LblCosto.Caption

                    
                    Agregando = False
                End If
            'Next A

            .Filter = adFilterNone
            .Filter = "corr = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & ""
            
            If chkAuto.Value = 0 Then
                Agregando = True
                Fg1.TextMatrix(Fg1.Row, 8) = RstRegistroSumar(RstPersonal, "canpro", "activo", "-1")
                Agregando = False
            End If
            
        End If
    End With
''    ' Se mandan a cargar los datos filtrados
''    pCargarDatos tareas, personal
End Sub

Private Sub hallarCosto(fgx As VSFlexGrid, columna As Integer)
    If fgx.Rows = fgx.FixedRows Then Exit Sub
    Dim A As Integer
    Dim costo As Double
    costo = 0
    For A = 1 To fgx.Rows - 2
        ' Si la celda esta seleccionada se aumenta el costo
        fgx.Select A, 2
        If fgx.CellChecked = flexChecked Then
            costo = costo + NulosN(fgx.TextMatrix(A, columna))
        End If
    Next A
    LblCosto.Caption = Format(NulosN(costo), "0.00000000")
End Sub

Private Sub pCargarDatos(Optional tareas As Boolean = True, Optional personal As Boolean = False)
    Agregando = True
    If tareas Then
        Fg3.Rows = 1
        With RstTar
            ' Se carga el detalle
            LblProd.Caption = NulosC(Fg1.TextMatrix(Fg1.Row, 5))
            lblIdRec.Caption = NulosN(Fg1.TextMatrix(Fg1.Row, 13))
            LblCosto.Caption = 0
            ' Se carga el contenido
            If .State = 0 Then Agregando = False: Exit Sub
            If .RecordCount = 0 Then Agregando = False: Exit Sub
            .MoveFirst
            Do While Not .EOF
                Fg3.Rows = Fg3.Rows + 1
                Fg3.TextMatrix(Fg3.Rows - 1, 1) = NulosN(.Fields("orden"))
                Fg3.TextMatrix(Fg3.Rows - 1, 2) = NulosC(.Fields("destar"))
                
                Fg3.Select Fg3.Rows - 1, 2
                If NulosN(.Fields("activo")) = True Then
                    Fg3.CellChecked = flexChecked
                    If NulosN(.Fields("costo")) <= 0 Then
                        Fg3.CellBackColor = vbRed
                    End If
                Else
                    Fg3.CellChecked = flexUnchecked
                End If
                
                Fg3.TextMatrix(Fg3.Rows - 1, 3) = Format(NulosN(.Fields("costo")), "0.00000000")
                Fg3.TextMatrix(Fg3.Rows - 1, 4) = NulosN(.Fields("idrec"))
                Fg3.TextMatrix(Fg3.Rows - 1, 5) = NulosN(.Fields("idpro"))
                Fg3.TextMatrix(Fg3.Rows - 1, 6) = NulosN(.Fields("idtar"))
                .MoveNext
            Loop
        End With
        ' aplicando el orden a la lista de datos
        GRID_ORDENAR Fg3, 1, 1
        hallarCosto Fg3, 3
    End If
    If personal Then
        Fg4.Rows = 1
        
        With RstPersonal
            If .State = 0 Then Agregando = False: Exit Sub
            If .RecordCount = 0 Then Agregando = False: Exit Sub
            .MoveFirst
            Do While Not .EOF
                Fg4.Rows = Fg4.Rows + 1
                Fg4.TextMatrix(Fg4.Rows - 1, 1) = NulosC(.Fields("codigo"))
                Fg4.TextMatrix(Fg4.Rows - 1, 2) = NulosC(.Fields("nombre"))
                
                Fg4.TextMatrix(Fg4.Rows - 1, 5) = Format(NulosN(.Fields("canpro")), FORMAT_CANTIDAD)
                
                Fg4.Select Fg4.Rows - 1, 2
                If NulosN(.Fields("activo")) = True Then
                    Fg4.TextMatrix(Fg4.Rows - 1, 3) = Format(.Fields("horini"), FORMAT_HORA_SIN_SEGUNDO)
                    Fg4.TextMatrix(Fg4.Rows - 1, 4) = Format(.Fields("horfin"), FORMAT_HORA_SIN_SEGUNDO)
                    
                    ' Se verifica si las horas de inicio y fin son vacias
                    If Format(.Fields("horini"), "HH:mm") = "00:00" _
                                        And .Fields("horini") = .Fields("horfin") Then
                        Fg4.TextMatrix(Fg4.Rows - 1, 3) = ""
                        Fg4.TextMatrix(Fg4.Rows - 1, 4) = ""
                    End If
                                        
                    Fg4.CellChecked = flexChecked
                Else
                    Fg4.CellChecked = flexUnchecked
                    '***********************************************
                    Fg4.TextMatrix(Fg4.Rows - 1, 3) = ""
                    Fg4.TextMatrix(Fg4.Rows - 1, 4) = ""
                    Fg4.TextMatrix(Fg4.Rows - 1, 5) = ""
                    '***********************************************
                End If
                
                Fg4.TextMatrix(Fg4.Rows - 1, 6) = NulosN(.Fields("idrec"))
                Fg4.TextMatrix(Fg4.Rows - 1, 7) = NulosN(.Fields("idper"))
                .MoveNext
            Loop
        End With
        ' aplicando el orden a la lista de datos
        '***********************************************
        'GRID_ORDENAR Fg4, 1, 2
        '***********************************************
        Fg4.Rows = Fg4.Rows + 1
    End If
    Agregando = False
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    
    If NulosN(Fg1.TextMatrix(Fg1.Row, 2)) <> 3 Then ' Si no es Linea
        If Fg1.Col <= 10 Or Fg1.Col = 18 Or Fg1.Col = 19 Or Fg1.Col = 20 Then
            Fg1.Editable = flexEDKbdMouse
        Else
            Fg1.Editable = flexEDNone
        End If
    Else
        If Fg1.Col = 3 Or Fg1.Col = 4 Then
            Fg1.Editable = flexEDNone
        Else
            Fg1.Editable = flexEDKbdMouse
        End If
    End If
End Sub

Private Sub Fg1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF6 Then
        If Fg1.Row < 1 Then Exit Sub
        If NulosN(Fg1.TextMatrix(Fg1.Row, 2)) = 1 Then Exit Sub
        If BuscarFrm("FrmControlTareaGr1", True, False) = True Then
            If FrmControlTareaGr1.fg(0).Rows > 1 Then
                FrmControlTareaGr1.fg(0).Row = 1
                FrmControlTareaGr1.fg(0).Col = 4
                FrmControlTareaGr1.fg(0).SetFocus
            Else
                FrmControlTareaGr1.Cmd(0).SetFocus
            End If
            Exit Sub
        End If
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    Select Case Col
        Case 3, 4, 5
            If validar_letras(KeyAscii) = False Then
                If validar_numero(KeyAscii) = False Then KeyAscii = 0
            End If
        
        Case 6, 7, 8, 20
           If validar_numero(KeyAscii) = False Then KeyAscii = 0
        
        Case 9     ' unidad
            KeyAscii = 0
        
        Case 1, 10 ' lote,comentario
        
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo error
    If KeyCode = 114 Or KeyCode = vbKeyInsert Then        'F3 = Agregar Item
        cmd_Click 0
    End If
    
    If KeyCode = 115 Or KeyCode = vbKeyDelete Then
        cmd_Click 1  'F4 = Eliminar Item
    End If
    Exit Sub
    
error:
    SHOW_ERROR Me.Name, "Fg1_KeyUp"
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace <> 3 Then
            PopupMenu Menu1
        End If
'        If QueHace = 3 Then
'            PopupMenu Menu4
'        Else
'            PopupMenu Menu1
'        End If
    End If
End Sub
 
Private Sub Fg1_RowColChange()
    If Agregando = True Then Exit Sub
    If Fg1.Rows = 1 Then Exit Sub
    
    If NulosN(Fg1.TextMatrix(Fg1.Row, 11)) = 0 Then Exit Sub    ' idtipo
    'If NulosN(Fg1.TextMatrix(Fg1.Row, 12)) = 0 And NulosN(Fg1.TextMatrix(Fg1.Row, 2)) <> 3 Then Exit Sub    ' idref
    
    If NulosN(Fg1.TextMatrix(Fg1.Row, 2)) = 1 Then  ' individual
        ' eliminar los datos registros del temporal
        RstRegistroEliminar RstGrDet, "codigo", NulosN(Fg1.TextMatrix(Fg1.Row, 16)), True
        Unload FrmControlTareaGr1
        
        '*********************
        Frm3.Visible = False
        Frm4.Visible = False
        '*********************
    Else
        If NulosN(Fg1.TextMatrix(Fg1.Row, 2)) <> 3 Then '--grupal
            ' mostrar en otra ventana los datos del grupo
            FrmControlTareaGr1.pRecibeLink Me.hWnd, NulosN(Fg1.TextMatrix(Fg1.Row, 16)), fActivarAutomaticoCantidad, fActivarAutomaticoHora
            FrmControlTareaGr1.Show
            If Fg1.Enabled = True Then Fg1.SetFocus
'            '--verificar si es lineal para mostrar el listado de tareas del producto
'            '--esta lista se mostrara por defecto activo a todas las tareas, el usuario desactivara si es necesario
'            If NulosN(Fg1.TextMatrix(Fg1.Row, 2)) = 3 Then
'                '--pendiente
'            End If
            '*********************
            Frm3.Visible = False
            Frm4.Visible = False
            '*********************
        Else '--lineal
            Dim filtro As String
            
'            If RstPersonal.State = 1 Then
'                If RstPersonal.RecordCount <> 0 Then
'                    RstPersonal.MoveFirst
'                    If NulosN(Fg1.TextMatrix(Fg1.Row, 16)) <> NulosN(RstPersonal("corr")) Then
'                        chkAuto.Value = 0
'                    End If
'                Else
'                    chkAuto.Value = 0
'                End If
'            End If
                        
            filtro = "corr = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & ""
        
            If RstTar.State <> 0 Then
                RstTar.Filter = filtro
                pCargarDatos True, False
            End If
            If RstPersonal.State <> 0 Then
                RstPersonal.Filter = filtro
                pCargarDatos False, True
            End If
                        
            '**********************************************************************************
            ' Se posiciona los frames para despejar el ingreso de tareas
            Dim valor As Integer
            Dim TOPE_FONDO As Double
            Dim TOPE_MENSAJE As Double
            Dim NUMERO_CELDAS_ As Integer
            
            TOPE_FONDO = TabOne1.Top + Frame2.Top + Fg1.Top
            
            If TOPE_FONDO < Frm3.Top Then TOPE_MENSAJE = Frm3.Top Else TOPE_MENSAJE = Frm3.Top + Frm3.Height
            
            valor = ((TOPE_MENSAJE - TOPE_FONDO) / Fg1.CellHeight) + Fg1.TopRow - 2
            valor = CInt(valor)
            
            If Fg1.Row > Abs(valor) Then
                Frm3.Top = 600
                Frm3.Left = Me.Width - 11800
                Frm4.Top = 200
                Frm4.Left = Me.Width - 6250
            Else
                Frm3.Top = Me.Height - 4150
                Frm3.Left = Me.Width - 11800
                Frm4.Top = Me.Height - 4150
                Frm4.Left = Me.Width - 6250
            End If
            '**********************************************************************************
                        
            Frm3.Visible = True
            cmdTar_click (0)
            Frm4.Visible = True
            Unload FrmControlTareaGr1
        End If
    End If
End Sub

Private Sub Fg3_Click()
    '*****************************************
    Dim columna As Integer
    Dim FILA As Integer
    Dim filaTop As Integer
    
    If QueHace = 3 Then Exit Sub
    
    filaTop = Fg3.TopRow
    
    FILA = Fg3.Row
    columna = Fg3.Col
    
    modificarRst RstTar, True, False ' se actualiza el recordset
    
    If FILA < Fg3.FixedRows Then Exit Sub
    If columna < Fg3.FixedCols Then Exit Sub
    
    If FILA > Fg3.Rows - 1 Then Exit Sub
    If columna > Fg3.Cols - 1 Then Exit Sub
    '****************************************************
    pCargarDatos True, False
    '****************************************************
    
    Fg3.TopRow = filaTop
    Fg3.Select FILA, columna
End Sub

Private Sub Fg3_KeyUp(KeyCode As Integer, Shift As Integer)
    '*****************************************
    Dim columna As Integer
    Dim FILA As Integer
    Dim filaTop As Integer
    
    If QueHace = 3 Then Exit Sub
    
    filaTop = Fg3.TopRow

    FILA = Fg3.Row
    columna = Fg3.Col

    modificarRst RstTar, True, False ' se actualiza el recordset
    
    If FILA < Fg3.FixedRows Then Exit Sub
    If columna < Fg3.FixedCols Then Exit Sub
    
    If FILA > Fg3.Rows - 1 Then Exit Sub
    If columna > Fg3.Cols - 1 Then Exit Sub

    Fg3.TopRow = filaTop
    Fg3.Select FILA, columna
End Sub

Private Sub Fg4_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim columna As Integer
    Dim FILA As Integer
    Dim filaTop As Integer
    
    If QueHace = 3 Then Exit Sub
    If Agregando = True Then Exit Sub
    
    If Col = 1 Then Exit Sub
    
    filaTop = Fg4.TopRow
    
    FILA = Fg4.Row
    columna = Fg4.Col
    
    If FILA < Fg4.FixedRows Then Exit Sub
    If columna < Fg4.FixedCols Then Exit Sub
    
    '--dando formato al grid
    Select Case Col
        Case 3, 4
            If IsDate(Fg4.TextMatrix(Row, Col)) = True Then
                Fg4.TextMatrix(Row, Col) = Format(Fg4.TextMatrix(Row, Col), FORMAT_HORA_SIN_SEGUNDO)
            Else
                Fg4.TextMatrix(Row, Col) = ""
            End If
        Case 5
            If IsNumeric(Fg4.TextMatrix(Row, Col)) = False Then
                Fg4.TextMatrix(Row, Col) = 0
            End If
            Fg4.TextMatrix(Row, Col) = Format(Fg4.TextMatrix(Row, Col), "0.00")
    End Select
    

    If Col = 2 Then
        modificarRst RstPersonal, False, True  ' se actualiza el recordset
    Else
        calcularCantidades
    End If
        
    
    If FILA > Fg4.Rows - 1 Then Exit Sub
    If columna > Fg4.Cols - 1 Then Exit Sub
    
    Fg4.TopRow = filaTop
    Fg4.Select FILA, columna
    
End Sub

Private Sub Fg4_Click()
    
    If Fg4.Col = 2 Then
        Fg4_CellChanged Fg4.Row, Fg4.Col
    End If
''    Dim columna As Integer
''    Dim FILA As Integer
''    Dim filaTop As Integer
''
''    If QueHace = 3 Then Exit Sub
''
''    filaTop = Fg4.TopRow
''
''    FILA = Fg4.Row
''    columna = Fg4.Col
''
''    If FILA < Fg4.FixedRows Then Exit Sub
''    If columna < Fg4.FixedCols Then Exit Sub
''
''    modificarRst RstPersonal, False, True ' se actualiza el recordset
''    calcularCantidades
''
''    If FILA > Fg4.Rows - 1 Then Exit Sub
''    If columna > Fg4.Cols - 1 Then Exit Sub
''
''    Fg4.TopRow = filaTop
''    Fg4.Select FILA, columna
End Sub

Private Sub Fg3_EnterCell()
    If QueHace = 3 Then
        Fg3.Editable = flexEDNone
        Exit Sub
    Else
        If Fg3.Col = 2 Then
            Fg3.Editable = flexEDKbdMouse
        Else
            Fg3.Editable = flexEDNone
        End If
    End If
End Sub

Private Sub Fg4_EnterCell()
    If QueHace = 3 Then
        Fg4.Editable = flexEDNone
        Exit Sub
    End If
'    Fg4.AllowUserResizing = flexResizeColumns
'    Fg4.ExplorerBar = flexExSortShowAndMove
'    Fg4.SelectionMode = flexSelectionFree
'
    Select Case Fg4.Col
        Case 1
        
        Case 2
            Fg4.AutoSearch = flexSearchFromCursor
            Fg4.AllowUserResizing = flexResizeColumns
            Fg4.ExplorerBar = flexExSortShowAndMove
            Fg4.SelectionMode = flexSelectionFree
        Case 3, 4, 5 '--Hora Inicio, Hora Fin, Cantidad
            Fg4.AutoSearch = flexSearchNone
            Fg4.Editable = flexEDKbdMouse
            Fg4.ExplorerBar = flexExNone
    End Select
End Sub


Private Sub Fg4_KeyUp(KeyCode As Integer, Shift As Integer)
    ' Detetecta el enter al activar o desactivar Personal
    'If KeyCode <> vbEnter Then Exit Sub
    Fg4_Click
    
'    Dim columna As Integer
'    Dim FILA As Integer
'    Dim filaTop As Integer
'
'    If QueHace = 3 Then Exit Sub
'
'    filaTop = Fg4.TopRow
'
'    FILA = Fg4.Row
'    columna = Fg4.Col
'
'    If FILA < Fg4.FixedRows Then Exit Sub
'    If columna < Fg4.FixedCols Then Exit Sub
'
'    modificarRst RstPersonal, False, True ' se actualiza el recordset
'    calcularCantidades
'
'    If FILA > Fg4.Rows - 1 Then Exit Sub
'    If columna > Fg4.Cols - 1 Then Exit Sub
'
'    Fg4.TopRow = filaTop
'    Fg4.Select FILA, columna
End Sub

Private Sub Fg4_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If QueHace = 3 Then Exit Sub
    
    Dim nSQL As String
    Dim nSQLId As String
    Dim nSQLTmp  As String
    Dim nTitulo As String
    Dim xform As New eps_librerias.FormSeleccion
    Dim xRs As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Cod. Empleado":        xCampos(0, 1) = "codemp":       xCampos(0, 2) = "2000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
    xCampos(1, 0) = "Apellidos y Nombres":  xCampos(1, 1) = "nombre":      xCampos(1, 2) = "5000":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
    xCampos(2, 0) = "Id":                   xCampos(2, 1) = "idemp":        xCampos(2, 2) = "1000":     xCampos(2, 3) = "C":    xCampos(2, 4) = "C"
        
    If Fg4.Rows = Fg4.FixedRows Then Fg4.Rows = Fg4.Rows + 1

    ' generar la lista de personal para no considerar en la lista
    RstPersonal.Filter = "corr = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & ""
    
    Dim COLSEL_ As Integer
    COLSEL_ = Fg4.Row
    
    nSQLId = GRID_GENERAR_SQL_ID(Fg4, 7, "AND pla_empleados.id", "NOT IN")
    
    If Fg4.Row <= 0 Then Fg4.Row = 1
    
    ' generar semejanzas en la lista
    If NulosC(Fg4.TextMatrix(Fg4.Row, 2)) <> "" Then
        nSQLTmp = " AND UCASE([pla_empleados].[nombre]) LIKE '%" & UCase(NulosC(Fg4.TextMatrix(Fg4.Row, 2))) & "%'"
    End If
    
    ' generar la consulta
    nSQL = "SELECT " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & " As corr, pla_empleados.codigo AS codemp, pla_empleados.id AS idemp, pla_empleados.nombre, -1 AS activo " _
        + vbCr + "FROM (pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper " _
        + vbCr + "Where (((pla_empleados.fchcese) Is Null) AND ((pro_empdet.idfun) = 6)) " & nSQLId & nSQLTmp _
        + vbCr + "ORDER BY pla_empleados.nombre;"
        
    nTitulo = "Buscando Personal"
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio
        
    xform.titulo = "Buscando Personal"
    
    If Not Agregando Then ' Si se esta modificando un registro
        eliminarRegistro
    End If
    
    If xRs.State = 1 Then
        ' agregando los datos al rst temporal
        RstPersonal.AddNew
        
        RstPersonal("corr") = NulosN(xRs("corr"))
        RstPersonal("idrec") = NulosN(lblIdRec.Caption)
        RstPersonal("idper") = xRs("idemp")
        RstPersonal("codigo") = NulosC(xRs("codemp"))
        RstPersonal("nombre") = NulosC(xRs("nombre"))
        RstPersonal("activo") = xRs("activo")
        RstPersonal("idunid") = 2
        RstPersonal("preuni") = NulosN(LblCosto.Caption)
        RstPersonal("imptot") = NulosN(LblCosto.Caption)
        
        RstPersonal("corr") = NulosN(xRs("corr"))
        RstPersonal("idrec") = NulosN(lblIdRec.Caption)
        RstPersonal("idper") = xRs("idemp")
        RstPersonal("codigo") = NulosC(xRs("codemp"))
        RstPersonal("nombre") = NulosC(xRs("nombre"))
        RstPersonal("activo") = xRs("activo")
        
        RstPersonal.Update
        
        If chkAuto.Value = 1 Then calcularCantidades
        
        pCargarDatos False, True
    
        seleccionarPersonal NulosN(xRs("idemp"))
    End If
    
    Agregando = False
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub seleccionarPersonal(IDPER_ As Double)
    If Fg4.Rows = Fg4.FixedRows Then Exit Sub
    Dim A As Integer
    Dim FILA_ As Double
    
    FILA_ = 1
    For A = 1 To Fg4.Rows - 1
        If NulosN(Fg4.TextMatrix(A, 7)) = IDPER_ Then FILA_ = A
    Next A
    
    Fg4.Select FILA_, 2
End Sub

Private Function GENERAR_SQL_ID2(Rst As ADODB.Recordset, nDesc As String, _
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
        
    GENERAR_SQL_ID2 = nSQL
End Function

Private Sub Fg4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '**************************************************************
    If Button = 2 Then
        If QueHace <> 3 Then
            PopupMenu Menu2
        End If
    End If
    '**************************************************************
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = True Then Exit Sub
   
    
    SeEjecuto = False
    mRowAdd = -999
    mRowAddTara = -9999
    mMesActivo = xMes
    
    '--Almacenar temporalmente el codigo del menu
    IdMenuActivo = xIdMenu
    
    pConfigurarGrilla
    
    pCargarGrid
    
    SeEjecuto = True
    If RstFrm.State = 0 Then Exit Sub
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
    QueHace = 3
    TabOne1.CurrTab = 0
    iniciarCampos
End Sub

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub

    If Me.Width <= 12000 Then Me.Width = 11985
    If Me.Height <= 8100 Then Me.Height = 8070

    ' Se dimensiona el Contenido
    TabOne1.Width = Me.Width - 100
    TabOne1.Height = Me.Height - 700
    
    Label4(0).Width = Me.Width - 100
    Dg3.Width = TabOne1.Width - 150
    Dg3.Height = TabOne1.Height - 1000
    
    lblperiodo(0).Left = TabOne1.Width - 1500
    
    ' Se dimensiona el Detalle
    Label1.Width = Me.Width - 100
    
    Frame4.Left = TabOne1.Width - 1850
    
    Fg1.Width = TabOne1.Width - 180
    Fg1.Height = TabOne1.Height - 2000
    
    Frame9.Top = TabOne1.Height - 1050
    Frame9.Width = TabOne1.Width - 100
    
    ' Se posiciona los frames para despejar el ingreso de tareas
    Dim valor As Integer
    Dim TOPE_FONDO As Double
    Dim TOPE_MENSAJE As Double
    Dim NUMERO_CELDAS_ As Integer
    
    TOPE_FONDO = TabOne1.Top + Frame2.Top + Fg1.Top
    
    If TOPE_FONDO < Frm3.Top Then TOPE_MENSAJE = Frm3.Top Else TOPE_MENSAJE = Frm3.Top + Frm3.Height
    
    valor = ((TOPE_MENSAJE - TOPE_FONDO) / Fg1.CellHeight) + Fg1.TopRow - 2
    valor = CInt(valor)
    
    If Fg1.Row > Abs(valor) Then
        Frm3.Top = 600
        Frm3.Left = Me.Width - 11800
        Frm4.Top = 200
        Frm4.Left = Me.Width - 6250
    Else
        Frm3.Top = Me.Height - 4150
        Frm3.Left = Me.Width - 11800
        Frm4.Top = Me.Height - 4150
        Frm4.Left = Me.Width - 6250
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    Else
        Set RstFrm = Nothing
        Set RstGrDet = Nothing
        Set RstGrDetTara = Nothing
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Menu1_1_Click
'* Tipo             : SUB
'* Descripcion      : Copia los registros seleccionados
'* Parametros       :
'* Devuelve         :
'* Creado por       : Jose Chacon 27/04/2011
'* Modificado       : 30/11/2011 - Copiar no solo los nombres sino tambien las Tareas
'*****************************************************************************************************
Private Sub Menu1_1_Click()
    Dim col_i As Integer
    Dim col_f As Integer
    Dim fil_i As Integer
    Dim fil_f As Integer
    Dim A As Integer
    Dim B As Integer
    Dim dato() As Variant
    
    With Fg1
        If .Col < .ColSel Then col_i = .Col: col_f = .ColSel
        If .Col > .ColSel Then col_f = .Col: col_i = .ColSel
        If .Col = .ColSel Then col_i = .Col: col_f = .Col
        
        If .Row < .RowSel Then fil_i = .Row: fil_f = .RowSel
        If .Row > .RowSel Then fil_f = .Row: fil_i = .RowSel
        If .Row = .RowSel Then fil_i = .Row: fil_f = .Row
        
        ReDim dato(col_i To col_f) As Variant
                
         For A = fil_i To fil_f
            For B = col_i To col_f
                dato(B) = NulosC(.TextMatrix(.Row, B))
                If B = 1 Then ' tipo persona
                    .TextMatrix(A, B) = dato(B)
                End If
                If B = 2 Then ' tipo persona
                    .TextMatrix(A, B) = dato(B)
                End If
                If B = 3 Then ' persona / grupo
                    .TextMatrix(A, B) = dato(B)
                    .TextMatrix(A, 12) = NulosN(.TextMatrix(.Row, 12))
                End If
                If B = 4 Then ' tarea
                    ' si la tarea es diferente => limpiar producto
                    If NulosN(.TextMatrix(A, 14)) <> 0 And (NulosN(.TextMatrix(A, 13)) <> NulosN(.TextMatrix(.Row, 13))) Then
                        .TextMatrix(A, 5) = "" ' Se limpia producto
                        .TextMatrix(A, 13) = "" ' se limpia id de producto
                    End If
                    .TextMatrix(A, B) = dato(B)
                    .TextMatrix(A, 14) = NulosN(.TextMatrix(.Row, 14))
                    ' Se copia la unidad por defecto
                    .TextMatrix(A, 9) = NulosC(.TextMatrix(.Row, 9))
                    .TextMatrix(A, 15) = NulosN(.TextMatrix(.Row, 15))
                End If
                If B = 5 Then ' producto
                    ' si la tarea es diferente => limpiar producto
                    If NulosN(.TextMatrix(A, 13)) <> 0 And (NulosN(.TextMatrix(A, 14)) <> NulosN(.TextMatrix(.Row, 14))) Then
                        .TextMatrix(A, 4) = "" ' Se limpia tarea
                        .TextMatrix(A, 14) = "" ' se limpia id de la tarea
                    End If
                    .TextMatrix(A, B) = dato(B)
                    .TextMatrix(A, 13) = NulosN(.TextMatrix(.Row, 13))
                    
                    
                    ' ****************************************************************
                    ' Se copian las tareas y personal solo para Linea
                    If NulosN(.TextMatrix(A, 2)) = 3 Then
                        If A = fil_i Then GoTo SALIRCOPIARPRODUCTO
                        
                        Dim IDRECETA_ As Double
                        Dim IDUNIDAD_ As Double
                        Dim CORRELATIVO_ As Double
                        Dim CORRELATIVOFUENTE_ As Double
                        Dim PRODUCTO_ As String
                        Dim RstAux As New ADODB.Recordset
                        
                        IDRECETA_ = NulosN(.TextMatrix(A, 13))
                        IDUNIDAD_ = NulosN(.TextMatrix(A, 15))
                        CORRELATIVO_ = NulosN(.TextMatrix(A, 16))
                        PRODUCTO_ = NulosC(.TextMatrix(A, 5))
                        Set RstAux = Nothing
                        
                        ' Se copian las Tareas
                        definir_llenar_Tareas IDRECETA_, IDUNIDAD_, CORRELATIVO_, PRODUCTO_
                        
                        ' Se copia el Personal
                        ' Se verifica la existencia del correlativo y se limpia
                        RstPersonal.Filter = "corr = " & CORRELATIVO_
                        If RstPersonal.RecordCount <> 0 Then limpiarRST RstPersonal, False
                        
                        RstPersonal.Filter = adFilterNone
                        CORRELATIVOFUENTE_ = NulosN(.TextMatrix(fil_i, 16))
                        RstPersonal.Filter = "corr = " & CORRELATIVOFUENTE_
                        
                        If RstPersonal.State = 0 Then GoTo SALIRCOPIARPRODUCTO
                        If RstPersonal.RecordCount = 0 Then GoTo SALIRCOPIARPRODUCTO
                        
                        DEFINIR_RST_TMP RstAux, RstPersonal
                        CARGAR_RST_TMP RstAux, RstPersonal
                        
                        RstAux.MoveFirst
                        While Not RstAux.EOF
                            RstPersonal.AddNew
                            RstPersonal("corr") = CORRELATIVO_
                            RstPersonal("idrec") = NulosN(RstAux("idrec"))
                            RstPersonal("idper") = NulosN(RstAux("idper"))
                            RstPersonal("codigo") = NulosC(RstAux("codigo"))
                            RstPersonal("nombre") = NulosC(RstAux("nombre"))
                            RstPersonal("activo") = NulosN(RstAux("activo"))
                            RstPersonal("idunid") = NulosN(RstAux("idunid"))
                            RstPersonal("preuni") = NulosN(RstAux("preuni"))
                            RstPersonal("imptot") = NulosN(RstAux("imptot"))
                            RstPersonal.Update
                            
                            RstAux.MoveNext
                        Wend
                    End If
                    Fg1_RowColChange
SALIRCOPIARPRODUCTO:
                    ' ****************************************************************
                    
                End If
                If B = 6 Then ' hora de inicio
                    .TextMatrix(A, B) = dato(B)
                End If
                If B = 7 Then ' hora de fin
                    .TextMatrix(A, B) = dato(B)
                End If
                If B = 8 Then ' cantidad
                    .TextMatrix(A, B) = dato(B)
                End If
                If B = 9 Then ' unidad de medida
                    .TextMatrix(A, B) = dato(B)
                    .TextMatrix(A, 15) = NulosN(.TextMatrix(.Row, 15))
                End If
            Next B
        Next A
    End With
End Sub

'*****************************************************************************************************
'* Nombre           : Menu1_2_Click
'* Tipo             : SUB
'* Descripcion      : Limpia los registros seleccionados
'* Parametros       :
'* Devuelve         :
'* Creado por       : Jose Chacon 27/04/2011
'* Modificado       :
'*****************************************************************************************************
Private Sub Menu1_2_Click()
    Dim col_i As Integer
    Dim col_f As Integer
    Dim fil_i As Integer
    Dim fil_f As Integer
    Dim A As Integer
    Dim B As Integer
    With Fg1
        If .Col < .ColSel Then col_i = .Col: col_f = .ColSel
        If .Col > .ColSel Then col_f = .Col: col_i = .ColSel
        If .Col = .ColSel Then col_i = .Col: col_f = .Col
        
        If .Row < .RowSel Then fil_i = .Row: fil_f = .RowSel
        If .Row > .RowSel Then fil_f = .Row: fil_i = .RowSel
        If .Row = .RowSel Then fil_i = .Row: fil_f = .Row
        
        For A = fil_i To fil_f
            For B = col_i To col_f
                .TextMatrix(A, B) = ""
                
                ' ****************************************************************
                If B = 5 Then ' Producto
                    ' Se copian las tareas y personal solo para Linea
                    If NulosN(.TextMatrix(A, 2)) = 3 Then
                        Dim IDRECETA_ As Double
                        Dim IDUNIDAD_ As Double
                        Dim CORRELATIVO_ As Double
                        Dim CORRELATIVOFUENTE_ As Double
                        Dim PRODUCTO_ As String
                        Dim RstAux As New ADODB.Recordset
                        
                        IDRECETA_ = NulosN(.TextMatrix(A, 13))
                        IDUNIDAD_ = NulosN(.TextMatrix(A, 15))
                        CORRELATIVO_ = NulosN(.TextMatrix(A, 16))
                        PRODUCTO_ = NulosC(.TextMatrix(A, 5))
                        Set RstAux = Nothing
                        
                        ' Se limpian las Tareas
                        RstTar.Filter = "corr = " & CORRELATIVO_
                        If RstTar.RecordCount <> 0 Then limpiarRST RstTar, False
                        
                        ' Se limpia el Personal
                        RstPersonal.Filter = "corr = " & CORRELATIVO_
                        If RstPersonal.RecordCount <> 0 Then limpiarRST RstPersonal, False
                    End If
                End If
                Fg1_RowColChange
                ' ****************************************************************
            Next B
        Next A
    End With
End Sub

'*****************************************************************************************************
'* Nombre           : Menu1_3_Click
'* Tipo             : SUB
'* Descripcion      : Elimina los registros seleccionados
'* Parametros       :
'* Devuelve         :
'* Creado por       : Jose Chacon 27/04/2011
'* Modificado       :
'*****************************************************************************************************
Private Sub Menu1_3_Click()
    If MsgBox("Esta seguro que desea eliminar los Registros seleccionados?", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    
    Dim fil_i As Integer
    Dim fil_f As Integer
    Dim A As Integer
    With Fg1
        If .Row < .RowSel Then fil_i = .Row: fil_f = .RowSel
        If .Row > .RowSel Then fil_f = .Row: fil_i = .RowSel
        If .Row = .RowSel Then fil_i = .Row: fil_f = .Row
        
        For A = fil_f To fil_i Step -1
            .RemoveItem A
        Next A
    End With
End Sub

Private Sub Menu2_1_Click()
    Dim col_i As Integer
    Dim col_f As Integer
    Dim fil_i As Integer
    Dim fil_f As Integer
    Dim A As Integer
    Dim B As Integer
    Dim dato() As Variant
    
    With Fg4
        ' Se calculan las filas de inicio y fin involucradas
        If .Col < .ColSel Then col_i = .Col: col_f = .ColSel
        If .Col > .ColSel Then col_f = .Col: col_i = .ColSel
        If .Col = .ColSel Then col_i = .Col: col_f = .Col
        
        If .Row < .RowSel Then fil_i = .Row: fil_f = .RowSel
        If .Row > .RowSel Then fil_f = .Row: fil_i = .RowSel
        If .Row = .RowSel Then fil_i = .Row: fil_f = .Row
        
        ReDim dato(col_i To col_f) As Variant
            
        Dim FILABASE_ As Integer ' Fila de la cual se copiara la informacion
        FILABASE_ = .Row
        For A = fil_i To fil_f
            Fg4.Select A, 2
            ' Se verifica si esta o no seleccionado el personal
            If Fg4.CellChecked = flexUnchecked Then GoTo SIGUIENTE
            
            For B = col_i To col_f
                dato(B) = NulosC(.TextMatrix(FILABASE_, B))
                Select Case B
                    Case 3, 4, 5 ' Hora de Inicio, Hora de Fin, Cantidad
                        .TextMatrix(A, B) = dato(B)
                        Fg4_CellChanged A, B
                End Select
            Next B
SIGUIENTE:
        Next A
        Fg4.Select Fg4.Row, col_f
    End With
End Sub

Private Sub menu2_2_Click()
    Dim col_i As Integer
    Dim col_f As Integer
    Dim fil_i As Integer
    Dim fil_f As Integer
    Dim A As Integer
    Dim B As Integer
    Dim dato() As Variant
    
    With Fg4
    ' Se calculan las hora de inicio y fin involucradas
        If .Col < .ColSel Then col_i = .Col: col_f = .ColSel
        If .Col > .ColSel Then col_f = .Col: col_i = .ColSel
        If .Col = .ColSel Then col_i = .Col: col_f = .Col
        
        If .Row < .RowSel Then fil_i = .Row: fil_f = .RowSel
        If .Row > .RowSel Then fil_f = .Row: fil_i = .RowSel
        If .Row = .RowSel Then fil_i = .Row: fil_f = .Row
                    
        Dim FILABASE_ As Integer
        FILABASE_ = .Row
        For A = fil_i To fil_f
            Fg4.Select A, 2
            If Fg4.CellChecked = flexUnchecked Then GoTo SIGUIENTE
            
            For B = col_i To col_f
                Select Case B
                    Case 3, 4, 5
                        .TextMatrix(A, B) = ""
                        Fg4_CellChanged A, B
                End Select
            Next B
SIGUIENTE:
        Next A
        Fg4.Select Fg4.Row, col_f
    End With
End Sub

Private Sub OptSeleccionar_Click(Index As Integer)
    If Index = 0 Then
        seleccionartodos
    End If
    
    If Index = 1 Then
        seleccionartodos False
    End If
    
    If Index = 2 Then
        seleccionartodos True, False, True
    End If
    
    If Index = 3 Then
        seleccionartodos False, False, True
    End If
End Sub

Private Sub seleccionartodos(Optional seleccionar As Boolean = True, Optional tareas As Boolean = True, _
                                    Optional personal As Boolean = False)
    Dim A As Integer
    
    If tareas Then
        If RstTar.State = 0 Then Exit Sub
        If RstTar.RecordCount <> 0 Then
            ' Agregando los datos al rst temporal
            RstTar.MoveFirst
            For A = 1 To Fg3.Rows - 1
                RstTar("activo") = seleccionar
                RstTar.MoveNext
                If RstTar.EOF = True Then Exit For
            Next A
        End If
    End If
    
    If personal Then
        If RstPersonal.State = 0 Then Exit Sub
        If RstPersonal.RecordCount <> 0 Then
            ' Agregando los datos al rst temporal
            RstPersonal.MoveFirst
            For A = 1 To Fg4.Rows - 1
                RstPersonal("activo") = seleccionar
                RstPersonal.MoveNext
                If RstPersonal.EOF = True Then Exit For
            Next A
        End If
    End If
    
    hallarCosto Fg3, 3
    calcularCantidades
    
    pCargarDatos tareas, personal
End Sub

Private Sub optTarea_Click(Index As Integer)
    txt_cb(3).Text = ""
    If Index = 0 Then
        lbl_cb_capt(3).Caption = "Tarea"
    Else
        lbl_cb_capt(3).Caption = "Receta"
    End If
End Sub

Private Sub PbCerrar_Click(Index As Integer)
    If Index = 0 Then Frm3.Visible = False: Frm4.Visible = False: Fg1.SetFocus
    If Index = 1 Then Frm4.Visible = False: Fg3.SetFocus
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    Unload FrmControlTareaGr1
    If OldTab = 0 Then
        If QueHace = 3 Then MuestraSegundoTab
    Else
        Frm3.Visible = False
        Frm4.Visible = False
        
        limpiarRST RstTar
        limpiarRST RstPersonal
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstFrm.Requery
            
            If RstFrm.RecordCount <> 0 Then
                RstFrm.MoveFirst
                RstFrm.Find "id=" & mIdRegistro
                If RstFrm.EOF = True Then RstFrm.MoveFirst
            End If
            
            Dg3.Refresh
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 8 Then Filtrar
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg3
        RstFrm.Filter = ""
    End If
    
    If Button.Index = 10 Then CambiarMes
    
    If Button.Index = 11 Then Buscar
    
    If Button.Index = 15 Then
        Set RstFrm = Nothing
        Unload Me
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
'    If ButtonMenu.Index = 1 Then pImprimir True
'
'    If ButtonMenu.Index = 2 Then pImprimir
    Dim ESTADOAUX_ As Double
    Dim MENSAJE_ As String
    Dim Rpta As Integer
    
    Select Case ButtonMenu.Parent.Index
        Case 2
            Select Case ButtonMenu.Index
                Case 1 ' Poner Pendiente Estado
                    ESTADOAUX_ = ESTADOPENDIENTE_
                    MENSAJE_ = "PENDIENTE"
                Case 2 ' Procesar Estado
                    ESTADOAUX_ = ESTADOPROCESADO_
                    MENSAJE_ = "PROCESADO"
                
                Case 3 ' Anular Estado
                    ESTADOAUX_ = ESTADOANULADO_
                    MENSAJE_ = "ANULADO"
                
            End Select
            
            Rpta = MsgBox("¿ Esta seguro de cambiar el estado actual a: " & MENSAJE_ & "?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbNo Then Exit Sub
            
            ' Se Actualiza el estado
            cSQL = "UPDATE pro_controltar SET pro_controltar.estado = " & ESTADOAUX_ & " " _
                    + vbCr + "WHERE (((pro_controltar.id)=" & NulosN(RstFrm("id")) & "));"
                    
            xCon.Execute cSQL
            
            ' grabamos el movimiento en la tabla var_edicion
            GrabarOperacion xIdUsuario, IdMenuActivo, 7, xHorIni, Time, Date, xCon, NulosN(RstFrm("id"))
            
            mIdRegistro = NulosN(RstFrm("id"))
            RstFrm.Requery
            Dg3.Refresh
            
            If RstFrm.RecordCount <> 0 Then
                RstFrm.MoveFirst
                RstFrm.Find "id=" & mIdRegistro
                If RstFrm.EOF = True Then RstFrm.MoveFirst
            End If
        Case 13
            Select Case ButtonMenu.Index
                Case 1
                    pImprimir True
                    
                Case 2
                    pImprimir
            End Select
    End Select
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA pro_controltar
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    On Error GoTo error
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        MsgBox "No hay registros", vbExclamation, xTitulo
        Exit Sub
    End If
    
    Dim xId&
    xId = NulosN(RstFrm.Fields("id"))
    TabOne1.CurrTab = 0
    If MsgBox("¿Esta seguro de eliminar el Seguimiento de las Tareas:" & vbCr & "Area: " & NulosC(RstFrm("area")) & "?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
    
        xCon.Execute "DELETe * FROM pro_controltardettar WHERE idctr = " & xId & ""
        xCon.Execute "DELETe * FROM pro_controltardetgrpes WHERE idctr = " & xId & ""
        xCon.Execute "DELETe * FROM pro_controltardetpes WHERE idctr = " & xId & ""
        xCon.Execute "DELETe * FROM pro_controltardetgr WHERE idctr = " & xId & ""
        xCon.Execute "DELETe * FROM pro_controltardet WHERE idctr = " & xId & ""
        xCon.Execute "DELETe * FROM pro_controltar WHERE id = " & xId & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & xId & " AND idform = " & IdMenuActivo
        
        MsgBox "El Seguimiento de la tarea :" & vbCr & "Area: " & NulosC(RstFrm("area")) & vbCr & "Dia:     " & Format(RstFrm("fchtra"), "dd/mm/yy") & vbCr & "Fue eliminado con éxito", vbInformation + vbOKOnly, xTitulo
        RstFrm.Requery
        Dg3.Refresh
        If RstFrm.RecordCount = 0 Then
            If MsgBox("No hay registrado ningúna producción, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
                Nuevo
            End If
        End If
    End If
    Exit Sub

error:
    SHOW_ERROR Me.Name, "Eliminar"
End Sub

'*****************************************************************************************************
'* Nombre           : iniciarCampos
'* Tipo             : SUB
'* Descripcion      :
'* Parametros       :
'* Devuelve         :
'* Creado por       : Jose Chacon
'* Modificado       : Jose Chacon 27/04/2011
'                       Se cambia Fg1.ExplorerBar = flexExSortShowAndMove por Fg1.ExplorerBar = flexExSortShow
'*****************************************************************************************************
Private Sub iniciarCampos()
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    Dg3.HeadLines = 2
    
    Fg1.AllowUserResizing = flexResizeColumns
    Fg1.AutoSearch = flexSearchFromTop
    Fg1.ExplorerBar = flexExSortShow
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.ForeColorSel = &H80000005
    Fg1.BackColorSel = &H80&
    
    Fg3.ColWidth(4) = 0
    Fg3.ColWidth(5) = 0
    Fg3.ColWidth(6) = 0
    
    Fg3.AllowUserResizing = flexResizeColumns
    Fg3.AutoSearch = flexSearchFromTop
    Fg3.ExplorerBar = flexExNone
    Fg3.SelectionMode = flexSelectionFree
    Fg3.ForeColorSel = &H80000005
    Fg3.BackColorSel = &H80&
    
    GRID_COMBOLIST Fg4, 2
    
    Fg4.ColWidth(6) = 0
    Fg4.ColWidth(7) = 0
    Fg4.ColWidth(8) = 0
    Fg4.ColWidth(9) = 0
    Fg4.ColWidth(10) = 0
    Fg4.ColWidth(11) = 0
    Fg4.ColWidth(12) = 0
        
    Fg4.AllowUserResizing = flexResizeColumns
    Fg4.AutoSearch = flexSearchFromTop
    Fg4.ExplorerBar = flexExSortShow
    Fg4.SelectionMode = flexSelectionFree
    Fg4.ForeColorSel = &H80000005
    Fg4.BackColorSel = &H80&
    
    ESTADOPENDIENTE_ = 1
    ESTADOPROCESADO_ = 2
    ESTADOANULADO_ = 6
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL PROCESO DE INGRESAR O MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    QueHace = 3
    TabOne1.TabEnabled(0) = True
    ActivaTool
    '**********************************
    Bloquea
    Frm3.Visible = False
    Frm4.Visible = False
    '**********************************
    pHabilitarObj False
    Label1.Caption = "Detalle del Seguimiento de Tareas"
    Fg1.SelectionMode = flexSelectionByRow
    Unload FrmControlTareaGr1
    TabOne1.CurrTab = 0
    Dg3.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA LA EDICION DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.EOF = True Or RstFrm.BOF = True Or RstFrm.RecordCount = 0 Then
        MsgBox "No hay Registros", vbExclamation, xTitulo
        Exit Sub
    End If

    QueHace = 2
    TabOne1.TabEnabled(0) = False
    ActivaTool
    '**********************************
    Bloquea
    
    Fg3.SelectionMode = flexSelectionFree
    Fg3.Editable = flexEDKbdMouse
    
    Fg4.SelectionMode = flexSelectionFree
    Fg4.Editable = flexEDKbdMouse
    '**********************************
    pHabilitarObj True
    
    If TabOne1.CurrTab = 0 Then
        TabOne1.CurrTab = 1
        MuestraSegundoTab
    End If
    
    Fg1.SelectionMode = flexSelectionFree
    Fg1.AutoSearch = flexSearchNone
    
    xHorIni = Time
    Label1.Caption = "Modificando Seguimiento de Tarea"
    txt_cb(1).SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EN FORMA DETALLADA EN LA PESTAÑA DETALLE DEL FORMULARIO LOS DATOS DE
'*                    UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub MuestraSegundoTab()
    iniciarCampos
    With RstFrm
        Blanquea
        If .State = 0 Then Exit Sub
        If .EOF = True Or .BOF = True Or .RecordCount = 0 Then Exit Sub
        If IsDate(.Fields("fchtra")) = True Then
            TxtFecha(0).valor = CDate(.Fields("fchtra"))
        End If
        
        txt_cb(0).Text = NulosN(RstFrm("idarea"))
        txt_cb_Validate 0, False
        lbl_cb(0).Caption = NulosC(RstFrm("area"))
        
        lbl_cod(0).Caption = NulosN(RstFrm("idarea"))
        
        txt_cb(1).Text = NulosC(RstFrm("numdoc"))
        lbl_cb(1).Caption = NulosC(RstFrm("encargado"))
        lbl_cod(1).Caption = NulosN(RstFrm("idres"))
        
        If NulosN(RstFrm("tipo")) = 1 Then
            OptTipoPago(0).Value = True ' pago x horas
        Else
            OptTipoPago(1).Value = True ' pago destajo
        End If
        
        llenarEstados Fg1, 21
        
        MuestraDetalle
    End With
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraDetalle
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EN FORMA DETALLADA LOS DATOS DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub MuestraDetalle()
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    
    ' limpiando el rst temporal
    Me.MousePointer = vbHourglass
    Set RstGrDet = Nothing
    DoEvents
    
'SELECT pro_controltardet.idctr, pro_controltardet.corr, pro_controltardet.numlote, pro_controltardet.tipo, IIf([pro_controltardet].[tipo]=1,[pla_empleados].[apepat] & ' ' & IIf([pla_empleados].[apemat] Is Null Or [pla_empleados].[apemat]='','',[pla_empleados].[apemat] & ' ') & [pla_empleados].[nom],'GRUPO Nº ' & [pro_controltardet].[idref]) AS nombres, alm_inventario.descripcion AS producto, pro_tareas.abrev AS tarea, pro_controltardet.horini, pro_controltardet.horfin, pro_controltardet.cant, mae_unidades.abrev, pro_controltardet.observacion, pro_controltardet.tipo AS idtipo, IIf([pro_controltardet].[tipo]=1,[pla_empleados].[id],[pro_controltardet].[idref]) AS idref, pro_controltardet.idrec, pro_controltardet.idtar, pro_controltardet.idunimed, pro_controltardet.observado, pro_controltardet.reproceso, pro_controltardet.cant1, pro_controltardet.estado, pro_controltardet.idprocorr, pro_producciondet.numparte AS numregprod
'FROM (pro_tareas RIGHT JOIN (mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN ((pro_receta RIGHT JOIN pro_controltardet ON pro_receta.id = pro_controltardet.idrec) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON alm_inventario.id = pro_receta.iditem) ON mae_unidades.id = pro_controltardet.idunimed) ON pro_tareas.id = pro_controltardet.idtar) LEFT JOIN pro_producciondet ON pro_controltardet.idprocorr = pro_producciondet.corr
'Where (((pro_controltardet.idctr) = 14064))
'ORDER BY pro_controltardet.tipo, IIf([pro_controltardet].[tipo]=1,[pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom],'GRUPO Nº ' & [pro_controltardet].[idref]), pro_controltardet.horini, alm_inventario.descripcion, pro_tareas.abrev;


    nSQL = "SELECT pro_controltardet.idctr, pro_controltardet.corr, pro_controltardet.numlote, pro_controltardet.tipo, IIf([pro_controltardet].[tipo]=1,[pla_empleados].[apepat] & ' ' & IIf([pla_empleados].[apemat] Is Null Or [pla_empleados].[apemat]='','',[pla_empleados].[apemat] & ' ') & [pla_empleados].[nom],'GRUPO Nº ' & [pro_controltardet].[idref]) AS nombres, alm_inventario.descripcion AS producto, pro_tareas.abrev AS tarea, pro_controltardet.horini, pro_controltardet.horfin, pro_controltardet.cant, mae_unidades.abrev, pro_controltardet.observacion, pro_controltardet.tipo AS idtipo, IIf([pro_controltardet].[tipo]=1,[pla_empleados].[id],[pro_controltardet].[idref]) AS idref, pro_controltardet.idrec, pro_controltardet.idtar, pro_controltardet.idunimed, pro_controltardet.observado, pro_controltardet.reproceso, pro_controltardet.cant1, pro_controltardet.estado, pro_controltardet.idprocorr, pro_producciondet.numparte AS numregprod " _
        + vbCr + "FROM (pro_tareas RIGHT JOIN (mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN ((pro_receta RIGHT JOIN pro_controltardet ON pro_receta.id = pro_controltardet.idrec) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON alm_inventario.id = pro_receta.iditem) ON mae_unidades.id = pro_controltardet.idunimed) ON pro_tareas.id = pro_controltardet.idtar) LEFT JOIN pro_producciondet ON pro_controltardet.idprocorr = pro_producciondet.corr " _
        + vbCr + " WHERE (((pro_controltardet.idctr)=" & RstFrm("id") & ")) " _
        + vbCr + " ORDER BY pro_controltardet.tipo, IIf([pro_controltardet].[tipo]=1,[pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom],'GRUPO Nº ' & [pro_controltardet].[idref]),pro_controltardet.horini, alm_inventario.descripcion, pro_tareas.abrev; "
    
'    nSQL = "SELECT pro_controltardet.idctr, pro_controltardet.corr, pro_controltardet.numlote, pro_controltardet.tipo, IIf([pro_controltardet].[tipo]=1,[pla_empleados].[apepat] & ' ' & iif([pla_empleados].[apemat] is null or [pla_empleados].[apemat]='','',[pla_empleados].[apemat] & ' ') & [pla_empleados].[nom],'GRUPO Nº ' & [pro_controltardet].[idref]) AS nombres, alm_inventario.descripcion AS producto, pro_tareas.abrev AS tarea, pro_controltardet.horini, pro_controltardet.horfin, pro_controltardet.cant, mae_unidades.abrev, pro_controltardet.observacion, pro_controltardet.tipo AS idtipo, IIf([pro_controltardet].[tipo]=1,[pla_empleados].[id],[pro_controltardet].[idref]) AS idref, pro_controltardet.idrec, pro_controltardet.idtar, pro_controltardet.idunimed,pro_controltardet.observado,pro_controltardet.reproceso,pro_controltardet.cant1 " _
'        + vbCr + " FROM pro_tareas RIGHT JOIN (mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN ((pro_receta RIGHT JOIN pro_controltardet ON pro_receta.id = pro_controltardet.idrec) LEFT JOIN pla_empleados ON pro_controltardet.idref = pla_empleados.id) ON alm_inventario.id = pro_receta.iditem) ON mae_unidades.id = pro_controltardet.idunimed) ON pro_tareas.id = pro_controltardet.idtar " _
'        + vbCr + " WHERE (((pro_controltardet.idctr)=" & RstFrm("id") & ")) " _
'        + vbCr + " ORDER BY pro_controltardet.tipo, IIf([pro_controltardet].[tipo]=1,[pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom],'GRUPO Nº ' & [pro_controltardet].[idref]),pro_controltardet.horini, alm_inventario.descripcion, pro_tareas.abrev; "

    RST_Busq RstTmp, nSQL, xCon
    DoEvents
    If RstTmp.RecordCount <> 0 Then
        DoEvents
        Agregando = True
        With Fg1
            .Rows = 1
            RstTmp.MoveFirst
            Do While Not RstTmp.EOF
                DoEvents
                .Rows = .Rows + 1
                '.TextMatrix(.Rows - 1, 1) = NulosC(RstTmp.Fields("numlote"))
                .TextMatrix(.Rows - 1, 2) = NulosN(RstTmp.Fields("tipo"))
                If NulosN(RstTmp.Fields("tipo")) <> 3 Then .TextMatrix(.Rows - 1, 3) = NulosC(RstTmp.Fields("nombres"))
                .TextMatrix(.Rows - 1, 4) = NulosC(RstTmp.Fields("tarea"))
                .TextMatrix(.Rows - 1, 5) = NulosC(RstTmp.Fields("producto"))
                If IsDate(RstTmp.Fields("horini")) = True Then .TextMatrix(.Rows - 1, 6) = Format(RstTmp.Fields("horini"), FORMAT_HORA_SIN_SEGUNDO)
                If IsDate(RstTmp.Fields("horfin")) = True Then .TextMatrix(.Rows - 1, 7) = Format(RstTmp.Fields("horfin"), FORMAT_HORA_SIN_SEGUNDO)
                .TextMatrix(.Rows - 1, 8) = NulosN(RstTmp.Fields("cant"))
                .TextMatrix(.Rows - 1, 9) = NulosC(RstTmp.Fields("abrev"))
                .TextMatrix(.Rows - 1, 10) = NulosC(RstTmp.Fields("observacion"))
                .TextMatrix(.Rows - 1, 11) = NulosN(RstTmp.Fields("idtipo"))
                If NulosN(RstTmp.Fields("tipo")) <> 3 Then .TextMatrix(.Rows - 1, 12) = NulosN(RstTmp.Fields("idref"))
                .TextMatrix(.Rows - 1, 13) = NulosN(RstTmp.Fields("idrec"))
                .TextMatrix(.Rows - 1, 14) = NulosN(RstTmp.Fields("idtar"))
                .TextMatrix(.Rows - 1, 15) = NulosN(RstTmp.Fields("idunimed"))
                .TextMatrix(.Rows - 1, 16) = NulosN(RstTmp.Fields("corr"))
                .TextMatrix(.Rows - 1, 17) = NulosN(RstTmp.Fields("cant"))
                .TextMatrix(.Rows - 1, 18) = NulosN(RstTmp.Fields("observado"))
                .TextMatrix(.Rows - 1, 19) = NulosN(RstTmp.Fields("reproceso"))
                .TextMatrix(.Rows - 1, 20) = NulosN(RstTmp.Fields("cant1"))
                
                '**************************************************************
                .TextMatrix(.Rows - 1, 1) = NulosC(RstTmp("numregprod"))
                .TextMatrix(.Rows - 1, 21) = NulosN(RstTmp("estado"))
                .TextMatrix(.Rows - 1, 22) = NulosN(RstTmp("idprocorr"))
                '**************************************************************
                
                RstTmp.MoveNext
            Loop
        End With
    End If
    
    ' cargar datos de los grupos
    pCargarDatosRstTemp 0, NulosN(RstFrm("id")), 0, True
    ' cargar datos de las taras vs peso
    pCargarDatosRstTemp 1, NulosN(RstFrm("id")), 0, True
        
    '******************************************************************
    ' cargar Tareas de linea
    pCargarDatosRstTemp 2, NulosN(RstFrm("id")), 0, True
    ' cargar grupos de linea
    pCargarDatosRstTemp 3, NulosN(RstFrm("id")), 0, True
    '******************************************************************
    
    Set RstTmp = Nothing
    GRID_AGRUPAR Fg1, 3
    Agregando = False
    Me.MousePointer = vbDefault
    Exit Sub

error:
    SHOW_ERROR Me.Name, "MuestraDetalle"
    Me.MousePointer = vbDefault
    Set RstTmp = Nothing
    Agregando = False
End Sub

'*****************************************************************************************************
'* Nombre           : pHabilitarObj
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA Y DESACTIVA LOS CONTROLES DEL FORMULARIO
'* Paranetros       : NOMBRE    |  TIPO         |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    band      |  Boolean      |
'* Devuelve         :
'*****************************************************************************************************
Private Sub pHabilitarObj(band As Boolean)
    habilitar_Locked TxtFecha, Not band
    habilitar_Locked txt_cb, Not band
    habilitar Me.cb, band
    habilitar Cmd, band
    Unload FrmControlTareaGr1
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA LOS CONTROLES TEXTBOX PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    LimpiaText TxtFecha
    LimpiaText txt
    LimpiaText txt_cb
    LimpiaText lbl_cod
    LimpiaText lbl_cb
    LimpiaText lblPesoTara
    
    '******************************************
    LimpiaText LblProd
    LimpiaText lblIdRec
    LimpiaText LblCosto
    Fg3.Rows = Fg3.FixedRows
    
    limpiarRST RstTar
    limpiarRST RstPersonal
    
    Frm3.Visible = False
    Frm4.Visible = False
    '******************************************
    
    Fg1.Rows = Fg1.FixedRows
    Set RstGrDet = Nothing
    Set RstGrDetTara = Nothing
    FraTarea.Visible = False
    FraEditor.Visible = False
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA Y DESCATIVA LOS BOTONES DE LA BARRA DE HERRAMIENTAS
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
'* Nombre           : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Nuevo()
    QueHace = 1
    mRowAdd = -999
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    '**********************************
    Bloquea
    
    Fg3.SelectionMode = flexSelectionFree
    Fg3.Editable = flexEDKbdMouse
    
    Fg4.SelectionMode = flexSelectionFree
    Fg4.Editable = flexEDKbdMouse
    '**********************************
    TxtFecha(0).valor = Date
    Blanquea
    pHabilitarObj True
    Label1.Caption = "Agregando Seguimiento de Tarea"
    TxtFecha(0).Enabled = True
    TxtFecha(0).SetFocus
    pConfigurarGrilla
    Fg1.SelectionMode = flexSelectionFree
    Fg1.AutoSearch = flexSearchNone
    
    '***********************
    llenarEstados Fg1, 21
    '***********************
    
    ' agregando un registro por defecto
    Fg1.Rows = 2
    Fg1.TextMatrix(Fg1.Rows - 1, 16) = mRowAdd       ' codigo de inicio
    Fg1.TextMatrix(Fg1.Rows - 1, 21) = ESTADOPROCESADO_
    ' cargar el temporal a la tara
    pCargarDatosRstTemp 1, -10, 0
    xHorIni = Time
    OptTipoPago(0).Value = True
    
End Sub

Sub Bloquea()
    OptSeleccionar(0).Enabled = Not OptSeleccionar(0).Enabled
    OptSeleccionar(1).Enabled = Not OptSeleccionar(1).Enabled
    
    OptSeleccionar(2).Enabled = Not OptSeleccionar(2).Enabled
    OptSeleccionar(3).Enabled = Not OptSeleccionar(3).Enabled
    
    cmdPer(0).Enabled = Not cmdPer(0).Enabled
    cmdPer(1).Enabled = Not cmdPer(1).Enabled
    cmdPer(2).Enabled = Not cmdPer(2).Enabled
    cmdPer(3).Enabled = Not cmdPer(3).Enabled
    cmdPer(4).Enabled = Not cmdPer(4).Enabled
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA pro_controltar, ESTA FUNCION DEVUELVE VERDADERO
'*                    CUANDO TIENE EXITO
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Function Grabar() As Boolean
'Modificado 21/01/11 Johan Castro
'           Agregar linea de codigo para grabar en campos de cabecera ano, idmes

    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " el Seguimiento de Tareas", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo SALIR
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstGr As New ADODB.Recordset
    Dim RstDetTara As New ADODB.Recordset
    Dim RstGrTara As New ADODB.Recordset
    Dim xId As Double
    Dim xCol&, xFil&, xItem&
    Dim HoraFraccion As Double
    Dim Difhora As String
    
    '******************************************
    Dim RstDetTareas As New ADODB.Recordset
    '******************************************

    On Error GoTo LaCague

    xCon.BeginTrans
    Me.MousePointer = vbHourglass
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT top 1 * FROM pro_controltar ", xCon
        xId = HallaCodigoTabla("pro_controltar", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstFrm("id")
        RST_Busq RstCab, "SELECT * FROM pro_controltar WHERE id =" & xId & "", xCon
        xCon.Execute "DELETE * FROM pro_controltardetgrpes WHERE idctr = " & xId & ""
        xCon.Execute "DELETE * FROM pro_controltardetgr WHERE idctr = " & xId & ""
        xCon.Execute "DELETE * FROM pro_controltardetpes WHERE idctr = " & xId & ""
        xCon.Execute "DELETE * FROM pro_controltardet WHERE idctr = " & xId & ""
        xCon.Execute "DELETE * FROM pro_controltardettar WHERE idctr = " & xId & ""
    End If
    
    mIdRegistro = xId
    RST_Busq RstDet, "SELECT top 1 * FROM pro_controltardet", xCon
    RST_Busq RstDetTara, "SELECT top 1 * FROM pro_controltardetpes", xCon
    RST_Busq RstGr, "SELECT top 1 * FROM pro_controltardetgr", xCon
    RST_Busq RstGrTara, "SELECT top 1 * FROM pro_controltardetgrpes", xCon
    RST_Busq RstDetTareas, "SELECT top 1 * FROM pro_controltardettar", xCon
    
    RstCab("fchtra") = CDate(TxtFecha(0).valor)
    RstCab("idarea") = NulosN(lbl_cod(0).Caption)
    RstCab("idres") = NulosN(lbl_cod(1).Caption)
    
    If OptTipoPago(0).Value = True Then
        RstCab("tipo") = 1
    Else
        RstCab("tipo") = 2
    End If
    
    RstCab("ano") = AnoTra
    RstCab("idmes") = mMesActivo
    
    '**************************************
    RstCab("estado") = ESTADOPROCESADO_
    '**************************************
    
    RstCab.Update
    
    For xFil = 1 To Fg1.Rows - 1
        RstDet.AddNew
        ' codigo
        RstDet("idctr") = xId
        RstDet("corr") = xFil
        RstDet("numlote") = NulosC(Fg1.TextMatrix(xFil, 1))
        RstDet("tipo") = NulosN(Fg1.TextMatrix(xFil, 11))
        RstDet("idref") = NulosN(Fg1.TextMatrix(xFil, 12))
        RstDet("idrec") = NulosN(Fg1.TextMatrix(xFil, 13))
        RstDet("idtar") = NulosN(Fg1.TextMatrix(xFil, 14))
        If IsDate(Fg1.TextMatrix(xFil, 6)) = True Then RstDet("horini") = CDate(Fg1.TextMatrix(xFil, 6))
        If IsDate(Fg1.TextMatrix(xFil, 7)) = True Then RstDet("horfin") = CDate(Fg1.TextMatrix(xFil, 7))
        RstDet("cant") = NulosN(Fg1.TextMatrix(xFil, 8))
        RstDet("idunimed") = NulosN(Fg1.TextMatrix(xFil, 15))
        RstDet("observacion") = NulosC(Fg1.TextMatrix(xFil, 10))
        RstDet("observado") = NulosN(Fg1.TextMatrix(xFil, 18))
        RstDet("reproceso") = NulosN(Fg1.TextMatrix(xFil, 19))
        RstDet("cant1") = NulosN(Fg1.TextMatrix(xFil, 20))
        
        If RstDet("tipo") = 1 Then
            ' calculando las horas de trabajo
            Difhora = DiferenciaHoras(Fg1.TextMatrix(xFil, 6), Fg1.TextMatrix(xFil, 7), True)
            HoraFraccion = Convert1HoraFaccion(Difhora)
            RstDet("tothor") = HoraFraccion
            If IsDate(Difhora) = True Then RstDet("difhor") = CDate(Difhora)
        End If
        
        '***************************************************
        RstDet("estado") = NulosN(Fg1.TextMatrix(xFil, 21))
        RstDet("idprocorr") = NulosN(Fg1.TextMatrix(xFil, 22))
        '***************************************************
        
        RstDet.Update
        ' registro de taras
        RstGrDetTara.Filter = "codigo= " & NulosN(Fg1.TextMatrix(xFil, 16)) & " and tipo=0 and idemp =" & NulosN(Fg1.TextMatrix(xFil, 16))
        If RstGrDetTara.RecordCount <> 0 Then RstGrDetTara.MoveFirst
        xItem = 1
        
        Do While Not RstGrDetTara.EOF
            RstDetTara.AddNew
            ' codigo
            RstDetTara("idctr") = xId
            RstDetTara("corr") = xFil
            RstDetTara("item") = xItem
            ' fin codigo
            RstDetTara("idpeso") = NulosN(RstGrDetTara.Fields("idpeso"))
            RstDetTara("pesouni") = NulosN(RstGrDetTara.Fields("pesouni"))
            RstDetTara("pesonet") = NulosN(RstGrDetTara.Fields("pesonet"))
            RstDetTara("cantidad") = NulosN(RstGrDetTara.Fields("cantidad"))
            RstDetTara("pesotara") = NulosN(RstGrDetTara.Fields("pesotara"))
            RstDetTara("pesobrut") = NulosN(RstGrDetTara.Fields("pesobrut"))
            RstDetTara.Update
            RstGrDetTara.MoveNext
            xItem = xItem + 1
        Loop
        
        ' grabar si es grupo
        If NulosN(Fg1.TextMatrix(xFil, 11)) = 2 Then
            RstGrDet.Filter = "codigo= " & NulosN(Fg1.TextMatrix(xFil, 16))
            If RstGrDet.RecordCount > 0 Then
                RstGrDet.MoveFirst
                Do While Not RstGrDet.EOF
                    RstGr.AddNew
                    ' codigo
                    RstGr("idctr") = xId
                    RstGr("corr") = xFil
                    RstGr("idper") = NulosN(RstGrDet.Fields("idemp"))
                    ' fin codigo
                    RstGr("cant") = NulosN(RstGrDet.Fields("cant"))
                    RstGr("cantbrut") = NulosN(RstGrDet.Fields("cantbrut"))
                    RstGr("activo") = NulosN(RstGrDet.Fields("activo"))
                    
                    If IsDate(RstGrDet.Fields("horini")) = True Then RstGr("horini") = CDate(RstGrDet.Fields("horini"))
                    If IsDate(RstGrDet.Fields("horfin")) = True Then RstGr("horfin") = CDate(RstGrDet.Fields("horfin"))
                    ' calculando las horas de trabajo
                    Difhora = DiferenciaHoras(NulosC(RstGrDet.Fields("horini")), NulosC(RstGrDet.Fields("horfin")), True)
                    HoraFraccion = Convert1HoraFaccion(Difhora)
                    RstGr("tothor") = HoraFraccion
                    If IsDate(Difhora) = True Then RstGr("difhor") = CDate(Difhora)
                    
                    RstGr.Update
                    
                    If NulosN(RstGrDet.Fields("activo")) = -1 Then     ' solo los activos
                        ' registro de taras
                        RstGrDetTara.Filter = "codigo= " & NulosN(Fg1.TextMatrix(xFil, 16)) & " and tipo=1 and idemp =" & NulosN(RstGrDet.Fields("idemp"))
                        If RstGrDetTara.RecordCount <> 0 Then RstGrDetTara.MoveFirst
                        xItem = 1
                        Do While Not RstGrDetTara.EOF
                            If NulosN(RstGrDetTara.Fields("pesobrut")) <> 0 And NulosN(RstGrDetTara.Fields("pesonet")) <> 0 Then
                                RstGrTara.AddNew
                                ' codigo
                                RstGrTara("idctr") = xId
                                RstGrTara("corr") = xFil
                                RstGrTara("idper") = NulosN(RstGrDet.Fields("idemp"))
                                RstGrTara("item") = xItem
                                ' fin codigo
                                RstGrTara("idpeso") = NulosN(RstGrDetTara.Fields("idpeso"))
                                RstGrTara("pesouni") = NulosN(RstGrDetTara.Fields("pesouni"))
                                RstGrTara("pesonet") = NulosN(RstGrDetTara.Fields("pesonet"))
                                RstGrTara("cantidad") = NulosN(RstGrDetTara.Fields("cantidad"))
                                RstGrTara("pesotara") = NulosN(RstGrDetTara.Fields("pesotara"))
                                RstGrTara("pesobrut") = NulosN(RstGrDetTara.Fields("pesobrut"))
                                RstGrTara.Update
                                xItem = xItem + 1
                            End If
                            RstGrDetTara.MoveNext
                        Loop
                    End If
                    RstGrDet.MoveNext
                Loop
            End If
        End If
        
        '*******************************************************************************************************
        ' grabar si es linea
        If NulosN(Fg1.TextMatrix(xFil, 11)) = 3 Then
            ' Se graba las tareas
            RstTar.Filter = "corr = " & NulosN(Fg1.TextMatrix(xFil, 16)) & ""
            If RstTar.RecordCount > 0 Then
                RstTar.MoveFirst
                Do While Not RstTar.EOF
                    RstDetTareas.AddNew
                    
                    RstDetTareas("idctr") = xId
                    RstDetTareas("corr") = xFil
                    RstDetTareas("idrec") = NulosN(RstTar.Fields("idrec"))
                    RstDetTareas("idtar") = NulosN(RstTar.Fields("idtar"))
                    RstDetTareas("orden") = NulosN(RstTar.Fields("orden"))
                    RstDetTareas("activo") = NulosN(RstTar.Fields("activo"))
                    
                    RstDetTareas.Update
                    
                    RstTar.MoveNext
                Loop
            End If
            
            ' Se graba al personal
            RstPersonal.Filter = "corr = " & NulosN(Fg1.TextMatrix(xFil, 16)) & ""
            If RstPersonal.RecordCount > 0 Then
                RstPersonal.MoveFirst
                Do While Not RstPersonal.EOF
                    RstGr.AddNew
                    RstGr("idctr") = xId
                    RstGr("corr") = xFil
                    RstGr("idper") = NulosN(RstPersonal.Fields("idper"))
                    RstGr("idrec") = NulosN(RstPersonal.Fields("idrec"))
                    RstGr("cant") = NulosN(RstPersonal.Fields("canpro"))
                    RstGr("canpro") = NulosN(RstPersonal.Fields("canpro"))
                    RstGr("cantbrut") = 0
                    RstGr("activo") = NulosN(RstPersonal.Fields("activo"))
                    
                    If IsDate(RstPersonal.Fields("horini")) = True Then RstGr("horini") = CDate(RstPersonal.Fields("horini"))
                    If IsDate(RstPersonal.Fields("horfin")) = True Then RstGr("horfin") = CDate(RstPersonal.Fields("horfin"))
                    ' calculando las horas de trabajo
                    Difhora = DiferenciaHoras(NulosC(RstPersonal.Fields("horini")), NulosC(RstPersonal.Fields("horfin")), True)
                    HoraFraccion = Convert1HoraFaccion(Difhora)
                    RstGr("tothor") = HoraFraccion
                    If IsDate(Difhora) = True Then RstGr("difhor") = CDate(Difhora)
                    
                    RstGr.Update
                    
                    RstPersonal.MoveNext
                Loop
            End If
        End If
        '*******************************************************************************************************
        
    Next xFil
    
    ' grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    xCon.CommitTrans
    
    MsgBox "El seguimiento de la Tarea se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    
    Grabar = True
    
SALIR:
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstGr = Nothing:      Set RstDetTareas = Nothing:
    Exit Function

LaCague:
    xCon.RollbackTrans
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstGr = Nothing:      Set RstDetTareas = Nothing:
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
    Grabar = False
End Function


'*****************************************************************************************************
'* Nombre           : fValidarDatos
'* Tipo             : FUNCION
'* Descripcion      : VERIFICA QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
'* Paranetros       :
'* Devuelve         : Boolean
'*****************************************************************************************************
Private Function fValidarDatos() As Boolean
    If TxtFecha(0).valor = "" Or IsDate(TxtFecha(0).valor) = False Then
        MsgBox "No ha especificado la fecha de la Programación ", vbInformation, xTitulo
        TxtFecha(0).SetFocus
        Exit Function
    End If
    
    Dim band As Integer
    band = Validar(txt_cb)
    
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado el registro de las tareas", vbInformation, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    
    ' Validar la grilla
    Dim mRow&, mCol&
    
    mCol = -1
    For mRow = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(mRow, 2)) = 0 Then       ' tipo de persona
            MsgBox "Seleccione el tipo Individual / Grupal / Lineal", vbExclamation, xTitulo
            mCol = 2
        '*************************************************************************************************************************
        ElseIf NulosN(Fg1.TextMatrix(mRow, 12)) = 0 And NulosN(Fg1.TextMatrix(mRow, 2)) <> 3 Then  ' persona/grupo
            MsgBox "Seleccione el " & IIf(NulosN(Fg1.TextMatrix(mRow, 2)) = 1, "Personal", "Nº de Grupo"), vbExclamation, xTitulo
            mCol = 3
        '*************************************************************************************************************************
        ElseIf NulosN(Fg1.TextMatrix(mRow, 15)) = 0 Then '--unidad de medida
            MsgBox "Falta ingresar la unidad de medida", vbExclamation, xTitulo
            mCol = 9
        End If

        If mCol <> -1 Then Exit For
    Next mRow
    
    If mCol <> -1 Then
        Agregando = True:  Fg1.Row = mRow: Fg1.Col = mCol: Agregando = False
        Fg1.SetFocus
        Exit Function
    End If

    fValidarDatos = True
End Function

'*****************************************************************************************************
'* Nombre           : pCargarGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS DATOS DE LA TABLA pro_controltar EN FUNCION A CRITERIOS ESPECIFICADOS
'*                    POR EL USUARIO EN EL CONTROL Dg3
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarGrid()
    On Error GoTo error
    Dim nSQL  As String
    
    lblperiodo(0).Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    lblperiodo(1).Caption = lblperiodo(0).Caption
    
    TDB_FiltroLimpiar Dg3
    Set RstFrm = Nothing
    
    '------------------------------------------------------------------------------------------
    '--bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    '------------------------------------------------------------------------------------------
    
'SELECT pro_controltar.id, pro_controltar.fchtra, pro_controltar.idarea, mae_area.descripcion AS area, pro_controltar.idres, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS encargado, pla_empleados.numdoc, pro_controltar.tipo, IIf(pro_controltar.tipo=1,'Horas','Destajo') AS TipoPago, pro_controltar.fchtra & '' AS fchtra1, UCase([mae_estados].[descripcion]) AS desestado
'FROM (mae_area RIGHT JOIN (pro_controltar LEFT JOIN (pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) ON pro_controltar.idres = pro_emp.id) ON mae_area.id = pro_controltar.idarea) LEFT JOIN mae_estados ON pro_controltar.estado = mae_estados.id
'Where (((pro_controltar.ano) = 2012) And ((pro_controltar.idmes) = 2))
'ORDER BY pro_controltar.fchtra DESC , mae_area.descripcion;

    nSQL = "SELECT pro_controltar.id, pro_controltar.fchtra, pro_controltar.idarea, mae_area.descripcion AS area, pro_controltar.idres, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS encargado, pla_empleados.numdoc, pro_controltar.tipo, IIf(pro_controltar.tipo=1,'Horas','Destajo') AS TipoPago, pro_controltar.fchtra & '' AS fchtra1, UCase([mae_estados].[descripcion]) AS desestado " _
        + vbCr + "FROM (mae_area RIGHT JOIN (pro_controltar LEFT JOIN (pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) ON pro_controltar.idres = pro_emp.id) ON mae_area.id = pro_controltar.idarea) LEFT JOIN mae_estados ON pro_controltar.estado = mae_estados.id " _
        + vbCr + "WHERE (([pro_controltar].[ano])=" & AnoTra & ") AND (([pro_controltar].[idmes]=" & mMesActivo & ")) " _
        + vbCr + "ORDER BY pro_controltar.fchtra desc,mae_area.descripcion "
    
    
'    nSQL = "SELECT pro_controltar.id, pro_controltar.fchtra, pro_controltar.idarea, mae_area.descripcion AS area, pro_controltar.idres, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS encargado,pla_empleados.numdoc,pro_controltar.tipo,iif(pro_controltar.tipo=1,'Horas','Destajo')  as TipoPago ,pro_controltar.fchtra & '' as fchtra1 " _
'        + vbCr + " FROM mae_area RIGHT JOIN (pro_controltar LEFT JOIN (pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) ON pro_controltar.idres = pro_emp.id) ON mae_area.id = pro_controltar.idarea " _
'        + vbCr + " WHERE (([pro_controltar].[ano])=" & AnoTra & ") AND (([pro_controltar].[idmes]=" & mMesActivo & ")) " _
'        + vbCr + " ORDER BY pro_controltar.fchtra desc,mae_area.descripcion "
        
    Me.MousePointer = vbHourglass
    RST_Busq RstFrm, nSQL, xCon
    
    Set Dg3.DataSource = RstFrm
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pCargarGrid"
End Sub

'*****************************************************************************************************
'* Nombre           : CambiarMes
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PERMITE CAMBIAR EL MES DE TRABAJO ACTUAL
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub CambiarMes()
    mMesActivo = SeleccionaMes(xCon)
    TabOne1.CurrTab = 0
    If mMesActivo = 0 Or mMesActivo = 13 Then
        MsgBox "Selecione un Periodo Correcto", vbExclamation, xTitulo
        CambiarMes
        Exit Sub
    End If
    pCargarGrid
End Sub

'*****************************************************************************************************
'* Nombre           : pImprimir
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPRIME LOS DATOS DEL CONTROL Dg3
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pImprimir(Optional IMP_LISTADO As Boolean = False)
    On Error GoTo error

    Me.MousePointer = vbHourglass
    If IMP_LISTADO = False Then
        If Me.TabOne1.CurrTab = 0 Then
        Else
''          MsgBox "Primero muestre el detalle del Registro" + vbCr + _
''              "Luego inténtelo otra vez", vbExclamation, xTitulo
        End If
    Else
        TDB_IMPRIMIR Dg3, "IMPRESIÓN DE PRODUCCIÓN", "LISTADO DE PRODUCCIÓN  -  Periodo: " + MonthName(mMesActivo, False)
    End If

    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"
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
    TabOne1.CurrTab = 0
     
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Fch.Trab":         xCampos(0, 1) = "fchtra":     xCampos(0, 2) = "1000":    xCampos(0, 3) = "F"
    xCampos(1, 0) = "Area":             xCampos(1, 1) = "area":       xCampos(1, 2) = "1500":   xCampos(1, 3) = "C"
    xCampos(2, 0) = "Respoinsable":     xCampos(2, 1) = "encargado":  xCampos(2, 2) = "3500":   xCampos(2, 3) = "C"
        
    nSQL = "SELECT pro_controltar.id, pro_controltar.fchtra, pro_controltar.idarea, mae_area.descripcion AS area, pro_controltar.idres, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS encargado " _
        + vbCr + " FROM mae_area RIGHT JOIN (pro_controltar LEFT JOIN (pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) ON pro_controltar.idres = pro_emp.id) ON mae_area.id = pro_controltar.idarea " _
        + vbCr + " WHERE (((Year([pro_controltar].[fchtra]))=" & AnoTra & ") AND ((Month([pro_controltar].[fchtra]))=" & mMesActivo & "));"

    CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), "Buscando Area", "fchtra", "fchtra", Principio
    If RstTmp.State = 0 Then GoTo SALIR
    If RstTmp.EOF = True And RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo SALIR
    
    RstFrm.MoveFirst
    RstFrm.Find "id = " + CStr(RstTmp("id"))

SALIR:
    Set RstTmp = Nothing
    Exit Sub

error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub

Private Sub llenarEstados(ByRef FGGRID As VSFlexGrid, columna As Integer)
    Dim CAMPOS As String
    Dim xRs As New ADODB.Recordset
    
    cSQL = "SELECT * FROM mae_estados ORDER BY id"
    RST_Busq xRs, cSQL, xCon
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then
        MsgBox "No se ha encontrado estados, Ingrese estados", vbInformation, xTitulo
        Exit Sub
    End If
    
    xRs.MoveFirst
    CAMPOS = "#" & NulosN(xRs("id")) & ";" & UCase(NulosC(xRs("descripcion")))
    xRs.MoveNext
    While Not xRs.EOF
        CAMPOS = CAMPOS & "|#" & NulosN(xRs("id")) & ";" & UCase(NulosC(xRs("descripcion")))
        xRs.MoveNext
    Wend
    FGGRID.ColComboList(columna) = CAMPOS
End Sub

'*****************************************************************************************************
'* Nombre           : Filtrar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA UN FILTRO AL RECORDSET RstFrm
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub Filtrar()
    Dim xCampos(2, 4) As String
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    xCampos(0, 0) = "Fch.Trab":         xCampos(0, 1) = "fchtra":     xCampos(0, 2) = "F":         xCampos(0, 3) = "800"
    xCampos(1, 0) = "Area":             xCampos(1, 1) = "area":       xCampos(1, 2) = "C":         xCampos(1, 3) = "1000"
    xCampos(2, 0) = "Respoinsable":     xCampos(2, 1) = "encargado":  xCampos(2, 2) = "C":         xCampos(2, 3) = "1500"
    TabOne1.CurrTab = 0
    CARGAR_DLL_EPSBUSCAR_FILTRO xCon, RstFrm, xCampos(), Dg3
End Sub

'*****************************************************************************************************
'* Nombre           : pConfigurarGrilla
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CONFIGURA EL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pConfigurarGrilla()
    Dim tFormat$
    Agregando = True
    With Fg1 '--de los ingredientes
        .Rows = 1
        .Cols = 23
        .FixedRows = 1
        .RowHeight(0) = 250
        .FrozenCols = 5
        .TextMatrix(0, 1) = "Nº Reg. Prod.":        .ColWidth(1) = 1200:    .ColAlignment(1) = flexAlignLeftCenter:   .Row = 0: .Col = 1:  .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 2) = "Tipo":                 .ColWidth(2) = 600:     .ColAlignment(2) = flexAlignLeftCenter:    .Row = 0: .Col = 2:  .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 3) = "Personal/Nº Grupo":    .ColWidth(3) = 1550:    .ColAlignment(3) = flexAlignLeftCenter:    .Row = 0: .Col = 3:  .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 4) = "Tarea":                .ColWidth(4) = 2000:    .ColAlignment(4) = flexAlignLeftCenter:    .Row = 0: .Col = 4:  .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 5) = "Producto":             .ColWidth(5) = 2400:    .ColAlignment(5) = flexAlignLeftCenter:    .Row = 0: .Col = 5:  .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 6) = "H.Inicio":             .ColWidth(6) = 800:     .ColAlignment(6) = flexAlignCenterCenter:  .Row = 0: .Col = 6:  .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 7) = "H.Final":              .ColWidth(7) = 800:     .ColAlignment(7) = flexAlignCenterCenter:  .Row = 0: .Col = 7:  .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 8) = "Cantidad":             .ColWidth(8) = 850:     .ColAlignment(8) = flexAlignRightCenter:   .Row = 0: .Col = 8:  .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 9) = "U.M.":                 .ColWidth(9) = 500:     .ColAlignment(9) = flexAlignCenterCenter:  .Row = 0: .Col = 9:  .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 10) = "Observacion":         .ColWidth(10) = 3000:   .ColAlignment(10) = flexAlignLeftCenter:   .Row = 0: .Col = 10: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 11) = "IdTipo":              .ColWidth(11) = 0:
        .TextMatrix(0, 12) = "IdRef":               .ColWidth(12) = 0:
        .TextMatrix(0, 13) = "IdRec":               .ColWidth(13) = 0:
        .TextMatrix(0, 14) = "IdTar":               .ColWidth(14) = 0:
        .TextMatrix(0, 15) = "IdUnimed":            .ColWidth(15) = 0:
        .TextMatrix(0, 16) = "Codigo":              .ColWidth(16) = 0:
        .TextMatrix(0, 17) = "CantTmp":             .ColWidth(17) = 0:     ' su uso sera para el calculo automatico
        .TextMatrix(0, 18) = "Obs":                 .ColWidth(18) = 0:      .ColAlignment(18) = flexAlignCenterCenter:  .Row = 0: .Col = 18: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 19) = "Reproceso":           .ColWidth(19) = 0:      .ColAlignment(18) = flexAlignCenterCenter:  .Row = 0: .Col = 19: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 20) = "Cant1":               .ColWidth(20) = 0:      .ColAlignment(20) = flexAlignRightCenter:   .Row = 0: .Col = 20: .CellAlignment = flexAlignCenterCenter
        
        '***************************************
        .TextMatrix(0, 21) = "Estado":              .ColWidth(21) = 1200:   .ColAlignment(21) = flexAlignLeftCenter:  .Row = 0: .Col = 21: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 22) = "idprocorr":           .ColWidth(22) = 0:
        '***************************************
        
        .ColEditMask(6) = "##:##"    ' hora inicio
        .ColEditMask(7) = "##:##"    ' hora fin
        .ColFormat(8) = FORMAT_MONTO ' cantidad
        
        .SelectionMode = flexSelectionByRow
        
        '**********************************************************
        GRID_COMBOLIST Fg1, 1        ' Numero de Reg. Produccion
        '**********************************************************
        
        GRID_COMBOLIST Fg1, 3        ' persona / grupo
        GRID_COMBOLIST Fg1, 4        ' tarea
        GRID_COMBOLIST Fg1, 5        ' producto
        GRID_COMBOLIST Fg1, 9        ' unidad de medida
        
        ' Tipo de Origen (Materia Prima; Producto)
        .ColComboList(2) = "#1;Individual|#2;Grupal|#3;Linea"
        .ColDataType(18) = flexDTBoolean
        .ColDataType(19) = flexDTBoolean
        DoEvents
    End With
    
    fg(1).ColWidth(6) = 0
    fg(1).ColWidth(7) = 0
    GRID_COMBOLIST fg(1), 3          ' peso - tara
    
    fg(0).ColWidth(3) = 0            ' idrec
    
    Agregando = False
    
    Fg4.ColEditMask(3) = "##:##"
    Fg4.ColEditMask(4) = "##:##"
End Sub

Private Sub cb_Click(Index As Integer)
    If QueHace = 3 Then Exit Sub
    Dim xCampos() As String
    Dim nCampoBusca As String
    Dim nSQL As String
    Dim nTitulo As String
    On Error GoTo error
    Select Case Index
        Case 0 ' area
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":           xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
            nTitulo = "Buscando Area"
            nSQL = "SELECT mae_area.id, mae_area.descripcion AS nombre, mae_area.id AS cod, mae_area.id AS idarea, pro_emp.id AS idper, pla_empleados.id AS idemp, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS encargado, pla_empleados.numdoc " _
                + vbCr + " FROM ((pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) RIGHT JOIN pro_area ON pro_emp.id = pro_area.idper) INNER JOIN mae_area ON pro_area.idarea = mae_area.id "
                
        Case 1 ' responsable de area
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":           xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
            nTitulo = "Buscando Responsable de Area "
            nSQL = "SELECT pla_empleados.numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, pro_emp.id AS cod, pla_empleados.id AS idemp, mae_dociden.abrev " _
                + vbCr + " FROM mae_dociden RIGHT JOIN ((pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) ON mae_dociden.id = pla_empleados.idtipdoc " _
                + vbCr + " WHERE (((pro_empdet.idfun)=5)); "
                
        Case 2 ' perdida de peso
            ReDim xCampos(3, 3) As String
            xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "2500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Peso":         xCampos(1, 1) = "peso":      xCampos(1, 2) = "900":    xCampos(1, 3) = "N"
            xCampos(2, 0) = "Und Destino":  xCampos(2, 1) = "destabrev": xCampos(2, 2) = "1500":   xCampos(2, 3) = "C"
            nTitulo = "Buscando Contenedor"
            nSQL = "SELECT pro_pesotara.id,  '1 ' & [mae_unidades].[abrev] & ' => ' & [pro_pesotara].[peso] & ' ' & [mae_unidades_1].[abrev] AS ref, pro_pesotara.id AS cod, pro_pesotara.descripcion as nombre,pro_pesotara.abrev, pro_pesotara.peso, mae_unidades_1.abrev AS destabrev  " _
                + vbCr + " FROM mae_unidades AS mae_unidades_1 INNER JOIN (mae_unidades INNER JOIN pro_pesotara ON mae_unidades.id = pro_pesotara.idundori) ON mae_unidades_1.id = pro_pesotara.idunddes;"

        Case 3 ' buscar las tareas relacionados a productos
            If OptTarea(0).Value = True Then
                ReDim xCampos(2, 3) As String
                xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
                xCampos(1, 0) = "Id":           xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
                nTitulo = "Buscando Tareas"
                nSQL = "SELECT pro_tareas.id, pro_tareas.descripcion as nombre, pro_tareas.id as cod FROM pro_tareas WHERE (((pro_tareas.diverso)=0)); "
            Else
                ReDim xCampos(3, 4) As String
                xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "codpro":   xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
                xCampos(1, 0) = "Descripción":  xCampos(1, 1) = "nombre":   xCampos(1, 2) = "5000":    xCampos(1, 3) = "C"
                xCampos(2, 0) = "CodReceta":    xCampos(2, 1) = "codrec":   xCampos(2, 2) = "1200":    xCampos(2, 3) = "C"
                nTitulo = "Buscando Recetas"
                nSQL = "SELECT pro_receta.codrec, alm_inventario.descripcion AS nombre, pro_receta.id AS cod, alm_inventario.codpro " _
                    + vbCr + " FROM alm_inventario INNER JOIN (pro_receta INNER JOIN pro_recetatar ON pro_receta.id = pro_recetatar.idrec) ON alm_inventario.id = pro_receta.iditem " _
                    + vbCr + " GROUP BY pro_receta.codrec, alm_inventario.descripcion, pro_receta.id,alm_inventario.codpro "
            End If
    End Select
    If Index <> 2 Then
    Else
    End If
    
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
        Case 0 ' area
            ' poner datos del encargado por defecto
            txt_cb(1).Text = NulosC(RstTmp.Fields("numdoc"))            ' TEXTO A MOSTRAR
            lbl_cb(1).Caption = NulosC(RstTmp.Fields("encargado"))      ' NOMBRE
            lbl_cod(1).Caption = NulosN(RstTmp.Fields("idper"))         ' CODIGO
            lbl_cb(1).ToolTipText = NulosC(RstTmp.Fields("encargado"))  ' NOMBRE
            If NulosN(lbl_cod(1).Caption) = 0 Then txt_cb(1).SetFocus
            If NulosN(lbl_cod(Index).Caption) = 4 Then
                Fg1.ColWidth(20) = 850
            Else
                Fg1.ColWidth(20) = 0
            End If
            
        Case 1 ' encargado
            If Fg1.Rows > Fg1.FixedRows Then
                Fg1.Col = 1:    Fg1.Row = Fg1.Rows - 1:   Fg1.SetFocus
            Else
                Cmd(0).SetFocus
            End If
            
        Case 2 ' perdida de peso
            lblPesoTara(0).Caption = NulosN(RstTmp("peso"))
            lblPesoTara(1).Caption = NulosC(RstTmp("abrev"))
            txt_cb(0).SetFocus
            
        Case 3 ' tareas relacionadas a productos
            If NulosN(lbl_cod(3).Caption) <> 0 Then pCargarTareasReceta
    
    End Select

SALIR:
    Set RstTmp = Nothing
    Exit Sub

error:
    Me.MousePointer = vbDefault
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If QueHace = 3 Then Exit Sub
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cod(Index).Caption = ""
    End If
    If Index = 0 Then txt_cb(1).Text = ""
    If Index = 2 Then LimpiaText lblPesoTara
    If Index = 3 Then fg(0).Rows = 1
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
    If QueHace = 3 Then Exit Sub
    If txt_cb(Index).Text = "" Then Exit Sub
    
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
        Case 0 ' area
            nSQL = "SELECT mae_area.id, mae_area.descripcion AS nombre, mae_area.id AS cod, mae_area.id AS idarea, pro_emp.id AS idper, pla_empleados.id AS idemp, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS encargado, pla_empleados.numdoc " _
                + vbCr + " FROM ((pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) RIGHT JOIN pro_area ON pro_emp.id = pro_area.idper) INNER JOIN mae_area ON pro_area.idarea = mae_area.id " _
                + vbCr + " WHERE mae_area.id = " & NulosN(txt_cb(Index).Text) & " ;"
        
        Case 1 ' encargado de area
            nSQL = "SELECT pla_empleados.numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, pro_emp.id AS cod, pla_empleados.id AS idemp, mae_dociden.abrev " _
                + vbCr + " FROM mae_dociden RIGHT JOIN ((pla_empleados RIGHT JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper) ON mae_dociden.id = pla_empleados.idtipdoc " _
                + vbCr + " WHERE (((pro_empdet.idfun)=5)) and pla_empleados.numdoc = '" & NulosC(txt_cb(Index).Text) & "';"
    
        Case 2 ' perdida de peso
            nSQL = "SELECT pro_pesotara.id,  '1 ' & [mae_unidades].[abrev] & ' => ' & [pro_pesotara].[peso] & ' ' & [mae_unidades_1].[abrev] AS ref, pro_pesotara.id AS cod, pro_pesotara.descripcion as nombre,pro_pesotara.abrev, pro_pesotara.peso, mae_unidades_1.abrev AS destabrev  " _
                + vbCr + " FROM mae_unidades AS mae_unidades_1 INNER JOIN (mae_unidades INNER JOIN pro_pesotara ON mae_unidades.id = pro_pesotara.idundori) ON mae_unidades_1.id = pro_pesotara.idunddes " _
                + vbCr + " WHERE pro_pesotara.id= " & NulosN(txt_cb(Index).Text) & ";"
        Case 3 ' tareas relacionadas con receta
            nSQL = "SELECT pro_tareas.id, pro_tareas.descripcion as nombre,pro_tareas.id as cod " _
            + vbCr + " FROM pro_tareas " _
            + vbCr + " WHERE pro_tareas.diverso=0 and pro_tareas.id= " & NulosN(txt_cb(Index).Text) & ";"
        
        Case Else
            Exit Sub
    End Select

    If xCon.State = 0 Then GoTo SALIR
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo SALIR

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cb(Index).Text = NulosC(RstTmp.Fields(0))                    ' TEXTO A MOSTRAR
        lbl_cb(Index).Caption = NulosC(RstTmp.Fields(1))                 ' NOMBRE
        lbl_cod(Index).Caption = NulosN(RstTmp.Fields(2))                ' CODIGO
        lbl_cb(Index).ToolTipText = NulosC(RstTmp.Fields(1))             ' NOMBRE
        If Index = 0 Then
            ' poner datos del encargado por defecto
            txt_cb(1).Text = NulosC(RstTmp.Fields("numdoc"))             ' TEXTO A MOSTRAR
            lbl_cb(1).Caption = NulosC(RstTmp.Fields("encargado"))       ' NOMBRE
            lbl_cod(1).Caption = NulosN(RstTmp.Fields("idper"))          ' CODIGO
            lbl_cb(1).ToolTipText = NulosC(RstTmp.Fields("encargado"))   ' NOMBRE
        ElseIf Index = 2 Then
            lblPesoTara(0).Caption = NulosN(RstTmp("peso"))
            lblPesoTara(1).Caption = NulosC(RstTmp("abrev"))
        End If
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cod(Index).Caption = ""
    End If
    
    If Index = 0 Then
'        If NulosN(lbl_cod(Index).Caption) = 4 Then
'            Fg1.ColWidth(20) = 850
'        Else
            Fg1.ColWidth(20) = 0
'        End If
    ElseIf Index = 3 Then
        pCargarTareasReceta
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

'*****************************************************************************************************
'* Nombre           : pRegistroAdd
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : AGREGA FILAS AL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pRegistroAdd()
    Dim mCol%
    Dim fInsertar As Boolean
    If QueHace = 3 Then Exit Sub
    Agregando = True
    If Fg1.Rows > Fg1.FixedRows Then
        If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 2)) = 0 Then         ' tipo de persona
            MsgBox "Seleccione el tipo Individual / Grupal", vbExclamation, xTitulo
            mCol = 2
        '*********************************************************************************************************************************
        ' Si es vacio y diferennte de lineal
        ElseIf NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 12)) = 0 And NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 2)) <> 3 Then    ' persona/grupo
            MsgBox "Seleccione el " & IIf(NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 2)) = 1, "Personal", "Nº de Grupo"), vbExclamation, xTitulo
            mCol = 3
        '*********************************************************************************************************************************
        ElseIf NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 15)) = 0 Then    ' unidad de medida
            MsgBox "Seleccione la unidad de medida", vbExclamation, xTitulo
            mCol = 9
        Else
            Frm3.Visible = False
            Frm4.Visible = False
            fInsertar = True
            mCol = 5
        End If
    Else
        fInsertar = True
        mCol = 1
    End If

    If fInsertar = True Then Fg1.AddItem ""
    
    If Fg1.Rows > 2 And fInsertar = True Then
        If chkOpcion(1).Value = 1 Then Fg1.TextMatrix(Fg1.Rows - 1, 1) = Fg1.TextMatrix(Fg1.Rows - 2, 1) ' num lote
        
        If chkOpcion(2).Value = 1 Then
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosN(Fg1.TextMatrix(Fg1.Rows - 2, 2))   ' tipo
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosN(Fg1.TextMatrix(Fg1.Rows - 2, 11)) ' idtipo
            mCol = 3
        End If
        
        If chkOpcion(3).Value = 1 Then
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(Fg1.TextMatrix(Fg1.Rows - 2, 3))   ' indivudual/grupal
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosN(Fg1.TextMatrix(Fg1.Rows - 2, 12)) ' cod ref dsfsf
            mCol = 4
        End If
        
        If chkOpcion(4).Value = 1 Then
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(Fg1.TextMatrix(Fg1.Rows - 2, 5))   ' tarea
            Fg1.TextMatrix(Fg1.Rows - 1, 14) = NulosN(Fg1.TextMatrix(Fg1.Rows - 2, 14)) ' idtar
            mCol = 5
        End If
        
        If chkOpcion(5).Value = 1 Then
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(Fg1.TextMatrix(Fg1.Rows - 2, 4))   ' producto
            Fg1.TextMatrix(Fg1.Rows - 1, 13) = NulosN(Fg1.TextMatrix(Fg1.Rows - 2, 13)) ' idrec
            mCol = 6
        End If
        
        If chkOpcion(6).Value = 1 Then
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(Fg1.TextMatrix(Fg1.Rows - 2, 6))   ' hora inicio
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosC(Fg1.TextMatrix(Fg1.Rows - 2, 7))   ' hora fin
            mCol = 8
        End If
        
        Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosC(Fg1.TextMatrix(Fg1.Rows - 2, 9))       ' unidad
        Fg1.TextMatrix(Fg1.Rows - 1, 15) = NulosN(Fg1.TextMatrix(Fg1.Rows - 2, 15))     ' idunid
        
        '******************************************************
        Fg1.TextMatrix(Fg1.Rows - 1, 21) = ESTADOPROCESADO_
        '******************************************************
        
        mRowAdd = mRowAdd + 1                      ' incrementar
        Fg1.TextMatrix(Fg1.Rows - 1, 16) = mRowAdd ' codigo
        
        If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 11)) = 2 Then
            pCargarDatosRstTemp 0, NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 12)), NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 16)), False
        End If
    End If
    
    Fg1.Row = Fg1.Rows - 1
    Fg1.Col = 4
    Agregando = False
    Fg1.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : pRegistroDel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA FILAS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pRegistroDel()
    If QueHace = 3 Then Exit Sub
    
    If Fg1.Row < 1 Then
        MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    If Fg1.Rows = 1 Then
        MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Sub
    End If
    If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    
    ' descargar el formulario de detalle del grupo
    Unload FrmControlTareaGr1
    ' eliminar los registros temporales
    RstRegistroEliminar RstGrDet, "codigo", NulosN(Fg1.TextMatrix(Fg1.Row, 16)), True
    
    '***************************************************************************************************
    RstRegistroEliminar RstPersonal, "corr", NulosN(Fg1.TextMatrix(Fg1.Row, 16)), True
    RstRegistroEliminar RstTar, "corr", NulosN(Fg1.TextMatrix(Fg1.Row, 16)), True
    
    Frm3.Visible = False
    Frm4.Visible = False
    '***************************************************************************************************
    
    Fg1.RemoveItem Fg1.Row
    If Fg1.Rows > 1 Then
        Fg1.Row = Fg1.Rows - 1
        Fg1.Col = 1
        Fg1.SetFocus
    Else
        Cmd(0).SetFocus
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarDatosRstTemp
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : Definir la estructura del recordset de los grupos
'* Paranetros       : NOMBRE       |   TIPO     |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    mTipo        |  Integer   |  Especifica el tipo de recordset a definir 0::Grupo;
'*                                                 1::Tara
'*                    idCod1       |            |  codigo del control de tareas
'*                    mRowPosicion |            |  posicion de la fila cuando se seleccione un grupo
'*                    fDesdeBD     |  Boolean   |  false::cuando se esta editando el grid
'*                                                 true::consulta directamente de la bd
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarDatosRstTemp(mTipo As Integer, idCod1, mRowPosicion, Optional fDesdeBD As Boolean = True)
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Set RstTmp = Nothing
    
    ' definir la estructura de recordset
    If RstGrDet.State = 0 Then

        nSQL = "SELECT pro_controltardetgr.corr AS codigo, pla_empleados.id AS idemp, pro_controltardet.idref AS idgrupo, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, pro_controltardetgr.cant, pro_controltardetgr.cantbrut, pro_controltardetgr.activo, pro_controltardetgr.horini, pro_controltardetgr.horfin  " _
            + vbCr + " FROM pla_empleados INNER JOIN (pro_controltardet INNER JOIN pro_controltardetgr ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) ON pla_empleados.id = pro_controltardetgr.idper " _
            + vbCr + " Where ((pro_controltardetgr.idctr) = -10)"
        RST_Busq RstTmp, nSQL, xCon
        DEFINIR_RST_TMP RstGrDet, RstTmp
    End If
    
    If mTipo = 0 Then
        ' cargar los datos
        If fDesdeBD = True Then
            nSQL = "SELECT pro_controltardetgr.corr AS codigo, pla_empleados.id AS idemp, pro_controltardet.idref AS idgrupo, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, pro_controltardetgr.cant, pro_controltardetgr.cantbrut, pro_controltardetgr.activo, pro_controltardetgr.horini, pro_controltardetgr.horfin " _
                + vbCr + " FROM pla_empleados INNER JOIN (pro_controltardet INNER JOIN pro_controltardetgr ON (pro_controltardet.corr = pro_controltardetgr.corr) AND (pro_controltardet.idctr = pro_controltardetgr.idctr)) ON pla_empleados.id = pro_controltardetgr.idper " _
                + vbCr + " WHERE ((pro_controltardetgr.idctr)=" & idCod1 & ") " _
                + vbCr + " ORDER BY pro_controltardetgr.corr, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom]; "
        Else
            nSQL = "SELECT " & mRowPosicion & " as codigo, pla_empleados.id AS idemp, pro_grupo.id AS idgrupo, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres,0 as cant,0 as cantbrut, -1 AS activo " _
                + vbCr + " FROM pro_grupo INNER JOIN ((pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_grupodet ON pro_emp.id = pro_grupodet.idper) ON pro_grupo.id = pro_grupodet.idgrupo " _
                + vbCr + " WHERE (((pro_grupo.id)=" & idCod1 & ")); "
        End If
    End If
    
    If mTipo = 1 Then
        nSQL = "select * from ( " _
            + vbCr + " SELECT pro_controltardetgr.corr AS codigo, 1 AS tipo, pro_controltardetgr.idper AS idemp, pro_controltardetgrpes.item, pro_pesotara.abrev, pro_controltardetgrpes.pesouni, pro_controltardetgrpes.pesonet, pro_controltardetgrpes.cantidad, pro_controltardetgrpes.pesotara, pro_controltardetgrpes.pesobrut, pro_controltardetgrpes.idpeso " _
            + vbCr + " FROM pro_pesotara RIGHT JOIN (pro_controltardetgr INNER JOIN pro_controltardetgrpes ON (pro_controltardetgr.idper = pro_controltardetgrpes.idper) AND (pro_controltardetgr.corr = pro_controltardetgrpes.corr) AND (pro_controltardetgr.idctr = pro_controltardetgrpes.idctr)) ON pro_pesotara.id = pro_controltardetgrpes.idpeso " _
            + vbCr + " Where (((pro_controltardetgr.idctr) = " & idCod1 & ")) " _
            + vbCr + " Union " _
            + vbCr + " SELECT pro_controltardet.corr AS codigo, 0 AS tipo, pro_controltardet.corr AS idemp, pro_controltardetpes.item, pro_pesotara.abrev, pro_controltardetpes.pesouni, pro_controltardetpes.pesonet, pro_controltardetpes.cantidad, pro_controltardetpes.pesotara, pro_controltardetpes.pesobrut, pro_controltardetpes.idpeso " _
            + vbCr + " FROM pro_controltardet INNER JOIN (pro_pesotara INNER JOIN pro_controltardetpes ON pro_pesotara.id = pro_controltardetpes.idpeso) ON (pro_controltardet.corr = pro_controltardetpes.corr) AND (pro_controltardet.idctr = pro_controltardetpes.idctr) " _
            + vbCr + " Where (((pro_controltardet.idctr) = " & idCod1 & ")) " _
            + vbCr + " ) as vw " _
            + vbCr + " order by vw.codigo, vw.tipo, vw.idemp, vw.item"
        RST_Busq RstTmp, nSQL, xCon
        If RstGrDetTara.State = 0 Then DEFINIR_RST_TMP RstGrDetTara, RstTmp
    End If
    
    '*************************************************************************************************************************************************************************************************
    If mTipo = 2 Then
        'Tareas
        nSQL = "SELECT vw_det.corr, vw_tareas.activo, vw_det.idrec, vw_det.idpro, vw_det.orden, vw_det.nompro, vw_det.codtar, vw_det.idtar, vw_det.destar, vw_costo.costo " _
            + vbCr + "FROM ( " _
            + vbCr + "( " _
            + vbCr + "SELECT DISTINCT pro_controltardet.idctr, pro_controltardet.corr, pro_receta.iditem AS idpro, pro_receta.id AS idrec, pro_recetatar.idtar, pro_controltardet.idunimed, pro_recetatar.orden, alm_inventario.descripcion AS nompro, pro_tareas.codigo AS codtar, pro_tareas.descripcion AS destar " _
            + vbCr + "FROM ((pro_receta INNER JOIN (pro_recetatar INNER JOIN pro_tareas ON pro_recetatar.idtar = pro_tareas.id) ON pro_receta.id = pro_recetatar.idrec) INNER JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) INNER JOIN pro_controltardet ON pro_recetatar.idrec = pro_controltardet.idrec " _
            + vbCr + "Where (((pro_controltardet.idctr) = " & idCod1 & ") And ((pro_controltardet.Tipo) = 3)) " _
            + vbCr + ") " _
            + vbCr + "AS vw_det Left Join " _
            + vbCr + "( " _
            + vbCr + "SELECT pro_controltardettar.idctr, pro_controltardettar.corr, pro_controltardettar.idrec, pro_controltardettar.idtar, pro_tareas.idunimed, pro_controltardettar.activo, pro_controltardettar.orden, pro_tareas.codigo AS codtar, pro_tareas.descripcion AS destar " _
            + vbCr + "FROM pro_controltardettar LEFT JOIN pro_tareas ON pro_controltardettar.idtar = pro_tareas.id " _
            + vbCr + "Where (((pro_controltardettar.idctr) = " & idCod1 & ") And ((pro_controltardettar.activo) = -1)) " _
            + vbCr + "ORDER BY pro_controltardettar.corr, pro_controltardettar.orden " _
            + vbCr + ") AS vw_tareas ON (vw_det.corr = vw_tareas.corr) AND (vw_det.idrec = vw_tareas.idrec) AND (vw_det.idtar = vw_tareas.idtar)) Left Join " _
            + vbCr + "( " _
            + vbCr + "SELECT pro_costo.idref AS idrec, pro_costodet.idtar, pro_costodet.idunimed, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.costo " _
            + vbCr + "FROM pro_tareas INNER JOIN (pro_costo INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
            + vbCr + "Where (((pro_costo.Tipo) = 1)) " _
            + vbCr + ") " _
            + vbCr + "AS vw_costo ON (vw_det.idunimed = vw_costo.idunimed) AND (vw_det.idtar = vw_costo.idtar) AND (vw_det.idrec = vw_costo.idrec);"
        
        
'        nSQL = "SELECT DISTINCT pro_controltardet.corr, IIf([tareas].[idrec] Is Null,0,-1) AS activo, pro_receta.id AS idrec, pro_receta.iditem AS idpro, pro_recetatar.orden, alm_inventario.descripcion AS nompro, pro_tareas.codigo AS codtar, pro_recetatar.idtar, pro_tareas.descripcion AS destar, costo.costo " _
'            + vbCr + "FROM ((((pro_receta INNER JOIN (pro_recetatar INNER JOIN pro_tareas ON pro_recetatar.idtar = pro_tareas.id) ON pro_receta.id = pro_recetatar.idrec) INNER JOIN alm_inventario ON pro_receta.iditem = alm_inventario.id) INNER JOIN pro_controltardet ON pro_recetatar.idrec = pro_controltardet.idrec) LEFT JOIN " _
'            + vbCr + "( " _
'            + vbCr + "SELECT pro_costo.idref, pro_costodet.idtar, pro_tareas.descripcion AS tarea, mae_unidades.abrev, pro_costodet.costo " _
'            + vbCr + "FROM pro_tareas INNER JOIN (pro_costo INNER JOIN (mae_unidades INNER JOIN pro_costodet ON mae_unidades.id = pro_costodet.idunimed) ON pro_costo.id = pro_costodet.idcos) ON pro_tareas.id = pro_costodet.idtar " _
'            + vbCr + "Where (((pro_costodet.idunimed) = 2) And ((pro_costo.Tipo) = 1)) " _
'            + vbCr + ") " _
'            + vbCr + "AS costo ON (pro_recetatar.idtar = costo.idtar) AND (pro_recetatar.idrec = costo.idref)) LEFT JOIN " _
'            + vbCr + "( " _
'            + vbCr + "SELECT  pro_controltardettar.corr, pro_controltardettar.idrec, pro_controltardettar.activo, pro_controltardettar.orden, pro_tareas.codigo AS codtar, pro_controltardettar.idtar, pro_tareas.descripcion AS destar, pro_controltardettar.idctr " _
'            + vbCr + "FROM pro_controltardettar LEFT JOIN pro_tareas ON pro_controltardettar.idtar = pro_tareas.id " _
'            + vbCr + "Where (((pro_controltardettar.activo) = -1) And ((pro_controltardettar.idctr) = " & idCod1 & ")) " _
'            + vbCr + "ORDER BY pro_controltardettar.corr, pro_controltardettar.orden " _
'            + vbCr + ") " _
'            + vbCr + "AS tareas ON (pro_recetatar.idtar = tareas.idtar) AND (pro_recetatar.idrec = tareas.idrec) " _
'            + vbCr + "WHERE (((pro_controltardet.tipo)=3) AND ((pro_controltardet.idctr)= " & idCod1 & "));"

        RST_Busq RstTmp, nSQL, xCon
        If RstTar.State = 0 Then DEFINIR_RST_TMP RstTar, RstTmp
    End If
    
    If mTipo = 3 Then
        'Personal
        nSQL = "SELECT pro_controltardetgr.corr, pro_controltardetgr.idrec, pro_controltardetgr.idper, pla_empleados.codigo, pla_empleados.nombre, pro_controltardetgr.activo, pro_controltardetgr.horini, pro_controltardetgr.horfin, pro_controltardetgr.tothor, pro_controltardetgr.difhor, pro_controltardetgr.canpro, pro_controltardetgr.idunid, pro_controltardetgr.preuni, pro_controltardetgr.imptot " _
            + vbCr + "FROM pro_controltardetgr LEFT JOIN pla_empleados ON pro_controltardetgr.idper = pla_empleados.id " _
            + vbCr + "WHERE (((pro_controltardetgr.idctr)= " & idCod1 & "));"
            
        RST_Busq RstTmp, nSQL, xCon
        If RstPersonal.State = 0 Then DEFINIR_RST_TMP RstPersonal, RstTmp
    End If
    '*************************************************************************************************************************************************************************************************
    
    Set RstTmp = Nothing
    RST_Busq RstTmp, nSQL, xCon
    
    DoEvents
    If mTipo = 0 Then
        If RstTmp.RecordCount <> 0 Then CARGAR_RST_TMP RstGrDet, RstTmp
    End If
    
    If mTipo = 1 Then
        If RstTmp.RecordCount <> 0 Then CARGAR_RST_TMP RstGrDetTara, RstTmp
    End If
    '*************************************************************************************************************************************************************************************************
    If mTipo = 2 Then 'Tareas
        If RstTmp.RecordCount <> 0 Then CARGAR_RST_TMP RstTar, RstTmp
    End If
    
    If mTipo = 3 Then 'Personal
        If RstTmp.RecordCount <> 0 Then CARGAR_RST_TMP RstPersonal, RstTmp
    End If
    '*************************************************************************************************************************************************************************************************
    
    Set RstTmp = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : pBuscarVSFlexGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA UNA BUSQUEDA EN EL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pBuscarVSFlexGrid()
    On Error GoTo error
    If Me.TabOne1.CurrTab = 0 Then Exit Sub
    Dim xExport As New SGI2_funciones.formularios
    Dim xCampos(3, 3) As String
    
    'campo     'columna del grid    'tipo(N:Numerico, C:caracter, F:fecha)      campo predeterminado(0:no se muestra, -1:se muestra al iniciar el formulario)
    xCampos(0, 0) = "Nº Orden":            xCampos(0, 1) = "1":    xCampos(0, 2) = "C":    xCampos(0, 3) = "0"
    xCampos(1, 0) = "Personal / Nº Grupo": xCampos(1, 1) = "3":    xCampos(1, 2) = "C":    xCampos(1, 3) = "-1"
    xCampos(2, 0) = "Producto":            xCampos(2, 1) = "5":    xCampos(2, 2) = "C":    xCampos(2, 3) = "0"
    xCampos(3, 0) = "Tarea":               xCampos(3, 1) = "4":    xCampos(3, 2) = "C":    xCampos(3, 3) = "0"
    
    xExport.VSFlexGrid_Buscar Me.hWnd, Fg1, xCampos(), Fg1.Row
    Set xExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pBuscarVSFlexGrid"
End Sub

'*****************************************************************************************************
'* Nombre           : pExportarVSFlexGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTAR A MS EXCEL LOS DATOS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pExportarVSFlexGrid()
    If IsDate(TxtFecha(0).valor) = False Then
        MsgBox "Falta especificar la Fecha de Trabajo", vbExclamation, xTitulo
        TxtFecha(0).SetFocus
        Exit Sub
    End If

    If lbl_cod(0).Caption = 0 Then
        MsgBox "Falta especificar el Area", vbExclamation, xTitulo
        txt_cb(0).SetFocus
        Exit Sub
    End If
    
    On Error GoTo error
    Dim oExport As New SGI2_funciones.formularios
    Dim nTitulo As String
    Dim nPeriodo As String
    Dim nTitulo1 As String
    nTitulo = "Control de Tareas - Area: " & StrConv(lbl_cb(0).Caption, 3)
    nPeriodo = "Fch. Trabajo: " + TxtFecha(0).valor
    nTitulo1 = "Responsable: " & StrConv(lbl_cb(1).Caption, 3)
    Me.MousePointer = vbHourglass
    oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg1, nTitulo, nPeriodo, nTitulo1, "Control de Tareas"
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
    
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportarVSFlexGrid"
End Sub

'*****************************************************************************************************
'* Nombre           : pHabilitarBotonEditor
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : Registrar las cantidades de ya se a una persona o a un grupo Ej. peso bruto y
'*                    cantidades
'*                    Mostrar/Ocultar las opciones del Ingreso de Datos
'* Paranetros       : NOMBRE    |  TIPO        |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Origen    |  Integer     |  =1::mostrar editor de taras-peso; 2::consultar tarea
'*                                                en recetas; 3::mostrar cuadro de opciones
'*                    band      |  Boolean     |  puede ser true o false
'* Devuelve         :
'*****************************************************************************************************
Private Sub pHabilitarBotonEditor(Origen As Integer, band As Boolean)
    Agregando = True
    habilitar Cmd, Not band
    habilitar CmdUtil, Not band
    Fg1.Enabled = Not band
    Toolbar1.Enabled = Not band
    habilitar cb, Not band
    habilitar txt_cb, Not band
    
    If Origen = 1 Then              ' fra_taras
        FraEditor.Visible = band
        ' true muestra el ingreso de datos
        If band = True Then
            If FrmControlTarea1.Fg1.Row <= 9 Then
                FraEditor.Top = 3535
                FraEditor.Left = 50
            Else
                FraEditor.Top = 265
                FraEditor.Left = 50
            End If
            LblTituloFrame.Caption = "Registros: " & Fg1.TextMatrix(Fg1.Row, 3)
        End If
        
        If band = True Then
            fg(1).Rows = 1
            With RstGrDetTara
                ' filtrar los registros solo del personal seleccionado
                .Filter = "codigo = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & " and idemp=" & NulosN(Fg1.TextMatrix(Fg1.Row, 16))
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not .EOF
                    fg(1).Rows = fg(1).Rows + 1
                    fg(1).TextMatrix(fg(1).Rows - 1, 1) = NulosN(.Fields("pesobrut"))
                    fg(1).TextMatrix(fg(1).Rows - 1, 2) = NulosN(.Fields("cantidad"))
                    fg(1).TextMatrix(fg(1).Rows - 1, 3) = NulosC(.Fields("abrev"))
                    fg(1).TextMatrix(fg(1).Rows - 1, 4) = NulosN(.Fields("pesouni"))
                    fg(1).TextMatrix(fg(1).Rows - 1, 5) = NulosN(.Fields("pesonet"))
                    fg(1).TextMatrix(fg(1).Rows - 1, 6) = NulosN(.Fields("idpeso"))
                    fg(1).TextMatrix(fg(1).Rows - 1, 7) = NulosN(.Fields("item"))       ' identificador de fila
                    .MoveNext
                Loop
            End With
            
            If fg(1).Rows > 1 Then
                fg(1).Row = fg(1).Rows - 1
                fg(1).Col = 1
                fg(1).SetFocus
            Else
                CmdEditor(0).SetFocus
            End If
            
            lblTotal(1).Caption = Format(GRID_SUMAR_COL(fg(1), 1), FORMAT_MONTO)
            lblTotal(3).Caption = Format(GRID_SUMAR_COL(fg(1), 5), FORMAT_MONTO)
        Else
            ' acumular las cantidades
            If NulosN(lblTotal(0).Caption) <> 0 Then Fg1.TextMatrix(Fg1.Row, 8) = GRID_SUMAR_COL(fg(1), 5)
            Fg1.Row = Fg1.Row
            Fg1.Col = 8
            Fg1.SetFocus
        End If
    ElseIf Origen = 2 Then ' fra_recetas(mostrar)
        FraTarea.Visible = band
        FraTarea.Enabled = band
        
        If band = True Then
            If FrmControlTarea1.Fg1.Row <= 9 Then
                FraTarea.Top = 2730
                FraTarea.Left = 50
            Else
                FraTarea.Top = 265
                FraTarea.Left = 50
            End If
            
            fg(0).Rows = 1
            fg(0).SelectionMode = flexSelectionByRow
            cb(3).Enabled = True
            txt_cb(3).Enabled = True
            txt_cb(3).Locked = False
            txt_cb(3).Text = ""
            txt_cb(3).SetFocus
        End If
    
    ElseIf Origen = 3 Then ' fra_recetas(mostrar)
        FraOpcion.Visible = band
        FraOpcion.Enabled = band
       
        If band = True Then
            FraOpcion.Top = 3990
            FraOpcion.Left = 4395
            chkOpcion(0).SetFocus
        End If
    End If
    Agregando = False
End Sub

Private Sub CmdEditor_Click(Index As Integer)
    Select Case Index
        Case 0 ' agregar
            pRegistroAddTara
        Case 1 ' eliminar
            pRegistroDelTara
        Case 2 ' cancelar
            pHabilitarBotonEditor 1, False
        Case 3 ' aceptar cuadro de opciones
            pHabilitarBotonEditor 3, False
    End Select
End Sub

Private Sub pic_Click(Index As Integer)
    If Index = 0 Then
        CmdEditor_Click 2
    ElseIf Index = 2 Then
        CmdTarea_Click 0
    ElseIf Index = 1 Then
        CmdEditor_Click 3
    End If
End Sub

Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If QueHace = 3 Then Exit Sub
    If Index <> 1 Then Exit Sub
    
    ' aplicando filtro
    RstGrDetTara.Filter = "codigo = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & " and tipo=0 and idemp = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16))
    If RstGrDetTara.RecordCount <> 0 Then RstGrDetTara.MoveFirst
    RstGrDetTara.Find "item = " & NulosN(fg(1).TextMatrix(Row, 7))
    
    If RstGrDetTara.EOF = False And RstGrDetTara.BOF = False Then
        RstGrDetTara("cantidad") = NulosN(fg(1).TextMatrix(Row, 2))
        RstGrDetTara("pesouni") = NulosN(fg(1).TextMatrix(Row, 4))
        RstGrDetTara("pesotara") = NulosN(RstGrDetTara("pesouni")) * NulosN(RstGrDetTara("cantidad"))
        RstGrDetTara("pesobrut") = NulosN(fg(1).TextMatrix(Row, 1))
        RstGrDetTara("pesonet") = NulosN(RstGrDetTara("pesobrut")) - NulosN(RstGrDetTara("pesotara"))
        fg(1).TextMatrix(Row, 5) = NulosN(RstGrDetTara("pesonet"))
    End If
    
    If IsNumeric(fg(1).TextMatrix(Row, Col)) = False Then fg(1).TextMatrix(Row, Col) = 0
    lblTotal(1).Caption = Format(GRID_SUMAR_COL(fg(1), 1), FORMAT_MONTO)
    lblTotal(3).Caption = Format(GRID_SUMAR_COL(fg(1), 5), FORMAT_MONTO)
    Fg1.TextMatrix(Fg1.Row, 8) = lblTotal(3).Caption
    fActivarAutomaticoCantidad = True
    Fg1_RowColChange
    fActivarAutomaticoCantidad = False
End Sub

Private Sub Fg_EnterCell(Index As Integer)
    If QueHace = 3 Or Index <> 1 Then
        fg(1).Editable = flexEDNone
        Exit Sub
    End If
    fg(1).Editable = flexEDKbdMouse
End Sub

Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Index <> 1 Then Exit Sub
    If Col <> 2 And Col <> 3 And Col <> 5 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLId As String
    Dim nTitulo As String
    
    If Col <> 3 Then Exit Sub
    ReDim xCampos(4, 3) As String
    xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombres": xCampos(0, 2) = "4500":  xCampos(0, 3) = "C"
    xCampos(1, 0) = "Peso":         xCampos(1, 1) = "peso":    xCampos(1, 2) = "800":   xCampos(1, 3) = "N"
    xCampos(2, 0) = "Abrev":        xCampos(2, 1) = "abrev":   xCampos(2, 2) = "700":   xCampos(2, 3) = "C"
    xCampos(3, 0) = "Id":           xCampos(3, 1) = "id":      xCampos(3, 2) = "500":   xCampos(3, 3) = "N"

    nTitulo = "Buscando Contenedor"
    
    nSQL = "SELECT pro_pesotara.id, pro_pesotara.descripcion AS nombres, pro_pesotara.peso, pro_pesotara.abrev " _
        + vbCr + " FROM mae_unidades INNER JOIN pro_pesotara ON mae_unidades.id = pro_pesotara.idundori "
        
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombres", "nombres", Principio, ""

    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    
    Agregando = True
    ' aplicando filtro
    RstGrDetTara.Filter = "codigo = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & " and tipo=0 and idemp=" & NulosN(Fg1.TextMatrix(Fg1.Row, 16))
    If RstGrDetTara.RecordCount <> 0 Then RstGrDetTara.MoveFirst
    RstGrDetTara.Find "item = " & NulosN(fg(1).TextMatrix(Row, 7))
    If RstGrDetTara.EOF = False And RstGrDetTara.BOF = False Then
        RstGrDetTara("abrev") = NulosC(xRs("abrev"))
        RstGrDetTara("idpeso") = NulosN(xRs("id"))
        RstGrDetTara("pesouni") = NulosN(xRs("peso"))
    End If
    
    ' actualizar el grid
    fg(1).TextMatrix(Row, 3) = NulosC(xRs("abrev"))
    fg(1).TextMatrix(Row, 6) = NulosN(xRs("id"))      ' idpeso
    fg(1).TextMatrix(Row, 4) = NulosN(xRs("peso"))
    
    Agregando = False
    fg_CellChanged 1, Row, 1
    fg(1).SetFocus
    Set xRs = Nothing
    Exit Sub
    
SALIR:
    Set xRs = Nothing
    Agregando = False
End Sub

Private Sub Fg_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> 1 Then Exit Sub
    If KeyCode = vbKeyEscape Then
        CmdEditor_Click 2
    ElseIf KeyCode = vbKeyF6 Then
        Fg1_KeyDown 117, 0
    End If
End Sub

Private Sub Fg_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo error
    If Index <> 1 Then Exit Sub
    If KeyCode = 114 Or KeyCode = vbKeyInsert Then CmdEditor_Click 0          ' F3 = Agregar Item
    If KeyCode = 115 Or KeyCode = vbKeyDelete Then CmdEditor_Click 1          ' F4 = Eliminar Item
    Exit Sub

error:
    SHOW_ERROR Me.Name, "Fg_KeyUp (" & Index & ")"
End Sub

'*****************************************************************************************************
'* Nombre           : pRegistroAddTara
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ADICIONA LAS TAREAS REALIZADAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pRegistroAddTara()
    ' 0 agregar personal; 1 agregar peso-taras
    Dim mCol%
    Dim fInsertar As Boolean
    
    Agregando = True
    fInsertar = True
    mCol = 1
    
    If fInsertar = True Then fg(1).AddItem ""
    
    fg(1).Row = fg(1).Rows - 1
    fg(1).Col = mCol
    ' cargar el buscador por defecto
    If fInsertar = True Then
        ' agregando el registro
        mRowAddTara = mRowAddTara + 1
        RstGrDetTara.AddNew
        RstGrDetTara("codigo") = NulosN(Fg1.TextMatrix(Fg1.Row, 16))
        RstGrDetTara("tipo") = 0
        RstGrDetTara("idemp") = NulosN(Fg1.TextMatrix(Fg1.Row, 16))
        RstGrDetTara("item") = mRowAddTara
        fg(1).TextMatrix(fg(1).Rows - 1, 7) = mRowAddTara
        
        ' colocar el ultimo peso-tara seleccionado
        If NulosN(RstGrDetTara("idpeso")) = 0 And fg(1).Rows > 2 Then
            RstGrDetTara("idpeso") = NulosN(fg(1).TextMatrix(fg(1).Rows - 2, 6))
            RstGrDetTara("pesouni") = NulosN(fg(1).TextMatrix(fg(1).Rows - 2, 4))
            RstGrDetTara("abrev") = NulosC(fg(1).TextMatrix(fg(1).Rows - 2, 3))
            RstGrDetTara("cantidad") = NulosN(fg(1).TextMatrix(fg(1).Rows - 2, 2))
        Else
            RstGrDetTara("idpeso") = NulosN(FrmControlTarea1.lbl_cod(2).Caption)
            RstGrDetTara("pesouni") = NulosN(FrmControlTarea1.lblPesoTara(0).Caption)
            RstGrDetTara("abrev") = NulosC(FrmControlTarea1.lblPesoTara(1).Caption)
            RstGrDetTara("cantidad") = 1
        End If
        
        fg(1).TextMatrix(fg(1).Row, 2) = NulosN(RstGrDetTara("cantidad"))
        fg(1).TextMatrix(fg(1).Row, 3) = NulosC(RstGrDetTara("abrev"))
        fg(1).TextMatrix(fg(1).Row, 4) = NulosN(RstGrDetTara("pesouni"))
        fg(1).TextMatrix(fg(1).Row, 6) = NulosN(RstGrDetTara("idpeso"))
    End If
    fg(1).SetFocus
    Agregando = False
End Sub

'*****************************************************************************************************
'* Nombre           : pRegistroDelTara
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UNA TAREA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pRegistroDelTara()
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
    RstGrDetTara.Filter = "codigo = " & NulosN(Fg1.TextMatrix(Fg1.Row, 16)) & " and tipo=0 and idemp=" & NulosN(Fg1.TextMatrix(Fg1.Row, 16))
    If RstGrDetTara.RecordCount <> 0 Then RstGrDetTara.MoveFirst
    RstGrDetTara.Find "item = " & NulosN(fg(1).TextMatrix(fg(1).Row, 7))
    If RstGrDetTara.EOF = False And RstGrDetTara.BOF = False Then
        RstGrDetTara.Delete
    End If
    
    fg(1).RemoveItem fg(1).Row
    If fg(1).Rows > 1 Then
        fg(1).Row = fg(1).Rows - 1
        fg(1).Col = 1
        fg(1).SetFocus
    Else
        CmdEditor(0).SetFocus
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarTareasReceta
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LAS TAREAS REGISTRADAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarTareasReceta()
    If NulosN(lbl_cod(3).Caption) = 0 Then
        txt_cb(3).SetFocus
        Exit Sub
    End If
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    
    On Error GoTo error
    
    Me.MousePointer = vbHourglass
    If OptTarea(0).Value = True Then
        nSQL = "SELECT alm_inventario.descripcion, pro_receta.codrec, pro_receta.id " _
            + vbCr + " FROM alm_inventario INNER JOIN (pro_tareas INNER JOIN (pro_receta INNER JOIN pro_recetatar ON pro_receta.id = pro_recetatar.idrec) ON pro_tareas.id = pro_recetatar.idtar) ON alm_inventario.id = pro_receta.iditem " _
            + vbCr + " WHERE (((pro_tareas.id)=" & NulosN(lbl_cod(3).Caption) & "))  " _
            + vbCr + " ORDER BY alm_inventario.descripcion, pro_receta.codrec;"
    Else
        nSQL = "SELECT pro_tareas.id, pro_tareas.descripcion, pro_recetatar.orden " _
            + vbCr + " FROM pro_tareas INNER JOIN pro_recetatar ON pro_tareas.id = pro_recetatar.idtar " _
            + vbCr + " Where (((pro_recetatar.idrec) =" & NulosN(lbl_cod(3).Caption) & ")) " _
            + vbCr + " ORDER BY pro_recetatar.orden;"
    End If

    RST_Busq RstTmp, nSQL, xCon
    DoEvents
    Agregando = True
    fg(0).ColWidth(3) = 0
    
    If OptTarea(0).Value = True Then
        fg(0).ColWidth(2) = 1020
        fg(0).TextMatrix(0, 1) = "Producto"
    Else
        fg(0).ColWidth(2) = 0
        fg(0).TextMatrix(0, 1) = "Tarea"
    End If
    
    If RstTmp.RecordCount <> 0 Then
        DoEvents
        With fg(0)
            .Rows = 1
            RstTmp.MoveFirst
            Do While Not RstTmp.EOF
                DoEvents
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = NulosC(RstTmp.Fields("descripcion"))
                If OptTarea(0).Value = True Then .TextMatrix(.Rows - 1, 2) = NulosC(RstTmp.Fields("codrec"))
                .TextMatrix(.Rows - 1, 3) = NulosN(RstTmp.Fields("id"))
                RstTmp.MoveNext
            Loop
        End With
    End If
    Set RstTmp = Nothing
    Agregando = False
    Me.MousePointer = vbDefault
    Exit Sub
    
error:
    Set RstTmp = Nothing
    Agregando = False
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pCargarTareasReceta"
End Sub

'*****************************************************************************************************
'* Nombre           : pConvertHora
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ESTABLECE EL TIEMPO TRANSCURRIDO ENTRE DOS HORAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pConvertHora(mRow As Long)
    If NulosN(Fg1.TextMatrix(mRow, 15)) = 7 Then
        Fg1.TextMatrix(mRow, 8) = ConvertHoraFaccion(Fg1.TextMatrix(mRow, 6), Fg1.TextMatrix(mRow, 7))
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : pImportar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPORTA DATOS DESDE UNA HOJA DE EXCEL
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pImportar()
    'Especificar las extensiones a usar
    Dim nPath As String
    Cmm.DefaultExt = "*.xls"
    Cmm.Filter = "Documentos de Excel (*.xls)|*.xls"
    Cmm.ShowOpen
    If Err Then
        'Cancelada la operación de abrir
    Else
        nPath = Cmm.FileName
    End If

    Dim A&
    Dim xNumFilas&
    Dim rstEmp As New ADODB.Recordset
    Dim RstTar As New ADODB.Recordset
    Dim RstProd As New ADODB.Recordset
    Dim nSQL As String
    Dim objExcel As Object
    
    Me.MousePointer = vbHourglass
    
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = True
    objExcel.SheetsInNewWorkbook = 1
    
    objExcel.WindowState = 2
    objExcel.Workbooks.Open nPath
    
    FraProgreso.Left = 3090
    FraProgreso.Top = 2910
    lbl(1).Caption = "Cargando registros para la importación"
    FraProgreso.Visible = True
    DoEvents
    
    xNumFilas = 1
    
    Fg1.Rows = 1
    With objExcel.ActiveSheet
        A = 2
        ' DETERMINAMOS EL NUMERO DE FILAS CON DATOS
        Do While NulosC(.Cells(A, 2)) <> ""
            If NulosC(.Cells(A, 2)) <> "" Then
                xNumFilas = xNumFilas + 1
            Else
                Exit Do
            End If
            A = A + 1
        Loop
        
        Fg1.Rows = 1
        xNumFilas = xNumFilas + 1
        PgBar.Max = xNumFilas
        
        ' cargando los datos para comparar personal
        nSQL = "SELECT pla_empleados.id AS idemp, [pla_empleados].[apepat] & ' ' & iif([pla_empleados].[apemat] is null or [pla_empleados].[apemat]='','',[pla_empleados].[apemat] & ' ') & [pla_empleados].[nom] AS personal, pla_empleados.fchnac " _
                + vbCr + " FROM (pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper " _
                + vbCr + " Where (((pro_empdet.idfun) = 6 )) " _
                + vbCr + " ORDER BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom]; "
        RST_Busq rstEmp, nSQL, xCon
       
        ' tarea
        nSQL = "SELECT pro_tareas.id, pro_tareas.descripcion AS tarea, pro_tareas.abrev FROM pro_tareas; "
        RST_Busq RstTar, nSQL, xCon
        
        ' producto
        nSQL = "SELECT DISTINCT alm_inventario.codpro, alm_inventario.descripcion AS producto, pro_receta.codrec, pro_receta.iditem, pro_receta.id AS idrec, pro_recetatar.idtar " _
                + vbCr + " FROM (alm_inventario INNER JOIN pro_receta ON alm_inventario.id = pro_receta.iditem) INNER JOIN pro_recetatar ON pro_receta.id = pro_recetatar.idrec; "
        RST_Busq RstProd, nSQL, xCon
        
        A = 2
        lbl(1).Caption = "Importando"
        Agregando = True
        DoEvents
        
        Do While NulosC(.Cells(A, 2)) <> ""
            PgBar.Value = A - 1
            DoEvents
            
            Fg1.Rows = Fg1.Rows + 1
            ' numlote
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(.Cells(A, 1))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = 1    ' tipo siempre indivual
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = 1
                ' persona
                rstEmp.Filter = ""
                rstEmp.Filter = "personal='" & NulosC(.Cells(A, 2)) & "'"
                If rstEmp.RecordCount <> 0 Then
                    Fg1.TextMatrix(Fg1.Rows - 1, 3) = .Cells(A, 2)
                    Fg1.TextMatrix(Fg1.Rows - 1, 12) = rstEmp("idemp")
                End If
                ' tarea
                RstTar.Filter = "tarea='" & NulosC(.Cells(A, 3)) & "' or abrev = '" & NulosC(.Cells(A, 3)) & "'"
                If RstTar.RecordCount <> 0 Then
                    Fg1.TextMatrix(Fg1.Rows - 1, 4) = .Cells(A, 3)
                    Fg1.TextMatrix(Fg1.Rows - 1, 14) = RstTar("id")
                    
                    ' producto
                    RstProd.Filter = "producto='" & NulosC(.Cells(A, 4)) & "' and idtar= " & RstTar("id")
                    If RstProd.RecordCount <> 0 Then
                        Fg1.TextMatrix(Fg1.Rows - 1, 5) = .Cells(A, 4)
                        Fg1.TextMatrix(Fg1.Rows - 1, 13) = RstProd("idrec")
                    End If
                End If
                Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(CDate(.Cells(A, 5)), FORMAT_HORA_SIN_SEGUNDO)
                Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(CDate(.Cells(A, 6)), FORMAT_HORA_SIN_SEGUNDO)
                
                ' cantidad
                Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosN(.Cells(A, 7))
                ' unidad de medida
                If NulosC(.Cells(A, 8)) <> "" Then
                    Fg1.TextMatrix(Fg1.Rows - 1, 15) = Busca_Codigo(NulosC(.Cells(A, 8)), "abrev", "id", "mae_unidades", "C", xCon)
                    If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 15)) <> 0 Then
                        Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosC(.Cells(A, 8))
                    End If
                End If
                ' observacion
                Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosC(.Cells(A, 9))
                mRowAdd = mRowAdd + 1
                Fg1.TextMatrix(Fg1.Rows - 1, 16) = mRowAdd
                ' si es por horas calcular el total de horas
                pConvertHora Fg1.Rows - 1
                DoEvents
            A = A + 1
        Loop
    End With
    
    FraProgreso.Visible = False
    Set RstTar = Nothing:    Set RstProd = Nothing:    Set rstEmp = Nothing
    Agregando = False
    
    MsgBox "El proceso termino de cargar los datos con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    objExcel.WindowState = 2
    objExcel.Workbooks.Close
    
    Set objExcel = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Set RstTar = Nothing:    Set RstProd = Nothing:    Set rstEmp = Nothing
    FraProgreso.Visible = False
    Me.MousePointer = vbDefault
    Set objExcel = Nothing
    SHOW_ERROR Me.Name, "pImportar"
End Sub

'Metodos para arrastrar el Frame
''''''''''''''''''''''''''''''''

Private Sub Frm4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    Frm4.ZOrder 0
End Sub

Private Sub Frm4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With Frm4
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub

Private Sub Frm3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    Frm3.ZOrder 0
End Sub

Private Sub Frm3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With Frm3
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub
