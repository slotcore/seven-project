VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsCosto 
   Caption         =   "Planillas - Consulta de Costo de Producción"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   615
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   11880
   Begin VB.CheckBox ChkGrupo 
      Caption         =   "Aplicar Grupo Horas/ Destajo/ Linea"
      Height          =   225
      Left            =   7890
      TabIndex        =   11
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CheckBox chkArea 
      Caption         =   "Aplicar Grupo x Area"
      Height          =   195
      Left            =   7890
      TabIndex        =   10
      Top             =   1080
      Width           =   2715
   End
   Begin VB.Frame Frame2 
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
      Height          =   1245
      Left            =   2130
      TabIndex        =   24
      Top             =   360
      Width           =   5625
      Begin VB.CommandButton cb 
         Height          =   240
         Index           =   2
         Left            =   1320
         Picture         =   "FrmConsCosto.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   900
         Width           =   225
      End
      Begin VB.CommandButton cb 
         Height          =   240
         Index           =   0
         Left            =   1320
         Picture         =   "FrmConsCosto.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   570
         Width           =   225
      End
      Begin VB.CommandButton cb 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "FrmConsCosto.frx":0264
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   225
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   1
         Left            =   870
         MaxLength       =   12
         TabIndex        =   2
         Text            =   "txt_cb(0)"
         Top             =   210
         Width           =   705
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   0
         Left            =   870
         MaxLength       =   12
         TabIndex        =   3
         Text            =   "txt_cb(0)"
         Top             =   540
         Width           =   705
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   2
         Left            =   870
         MaxLength       =   12
         TabIndex        =   4
         Text            =   "txt_cb(2)"
         Top             =   870
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Personal"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Area"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   645
         Width           =   330
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "T. Planilla"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   330
         Width           =   690
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
         Left            =   2670
         TabIndex        =   32
         Top             =   870
         Visible         =   0   'False
         Width           =   1005
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
         Left            =   2670
         TabIndex        =   29
         Top             =   540
         Visible         =   0   'False
         Width           =   1005
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
         Left            =   1590
         TabIndex        =   27
         Top             =   210
         Width           =   2295
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
         Left            =   2670
         TabIndex        =   26
         Top             =   210
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
         Left            =   1590
         TabIndex        =   30
         Top             =   540
         Width           =   2295
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
         Left            =   1590
         TabIndex        =   33
         Top             =   870
         Width           =   3915
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Turno"
      Height          =   615
      Left            =   10380
      TabIndex        =   23
      Top             =   360
      Width           =   1455
      Begin VB.CheckBox chk 
         Caption         =   "Noche"
         Height          =   225
         Index           =   3
         Left            =   630
         TabIndex        =   9
         Top             =   240
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.CheckBox chk 
         Caption         =   "Dia"
         Height          =   225
         Index           =   2
         Left            =   60
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Checked
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   615
      Left            =   7770
      TabIndex        =   22
      Top             =   360
      Width           =   2595
      Begin VB.CheckBox chk 
         Caption         =   "Linea"
         Height          =   225
         Index           =   4
         Left            =   1770
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   705
      End
      Begin VB.CheckBox chk 
         Caption         =   "Horas"
         Height          =   225
         Index           =   1
         Left            =   950
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chk 
         Caption         =   "Destajo"
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
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
      Height          =   1245
      Left            =   30
      TabIndex        =   19
      Top             =   360
      Width           =   2085
      Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
         Height          =   300
         Index           =   0
         Left            =   540
         TabIndex        =   0
         Top             =   300
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
         Valor           =   "25/09/2007"
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
         Height          =   300
         Index           =   1
         Left            =   540
         TabIndex        =   1
         Top             =   780
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
         Valor           =   "25/09/2007"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "A:"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   870
         Width           =   150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "De:"
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   405
         Width           =   255
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
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
               Picture         =   "FrmConsCosto.frx":0396
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCosto.frx":08DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCosto.frx":0C6C
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCosto.frx":0DF0
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCosto.frx":1244
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCosto.frx":135C
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCosto.frx":18A0
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCosto.frx":1DE4
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCosto.frx":1EF8
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCosto.frx":200C
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCosto.frx":2460
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCosto.frx":25CC
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCosto.frx":2B14
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsCosto.frx":2E2E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraProgreso 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   3840
      TabIndex        =   13
      Top             =   3570
      Visible         =   0   'False
      Width           =   5760
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   90
         TabIndex        =   14
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
         Caption         =   "Registros"
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
         TabIndex        =   17
         Top             =   75
         Width           =   810
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
         TabIndex        =   16
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
         TabIndex        =   15
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
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   5865
      Left            =   30
      TabIndex        =   18
      Top             =   1650
      Width           =   11835
      _cx             =   20876
      _cy             =   10345
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
      BackColor       =   14745342
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   128
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14745342
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
      FormatString    =   $"FrmConsCosto.frx":31C0
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
Attribute VB_Name = "FrmConsCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMCONSCOSTO.FRM
'* Tipo             : FORMULARIO
'* Descripcion      : FORMULARIO PARA CONSULTAR EL PAGO QUE SE REALIZARA A CADA TRABAJADOR EN FUNCION
'*                    A CRITERIOS ESPECIFICADOS POR EL USUARIO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 28/10/09
'* VERSION          : 1.0
'*****************************************************************************************************

'--HISTORIA
'--Modificado  25/05/10 por Johan Castro
'--            Agregar columna para mostrar el numero de documento de personal
'--Modificado  12/11/11 por Jose Chacon
'--            Agregar Filtro Tipo de Planilla

Option Explicit

Dim BAND_INTERRUMPIR As Boolean           ' SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                          ' TRUE SE INTERRUMPE
Dim SeEjecuto  As Boolean                 ' VERIFICA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLA VEZ

'*****************************************************************************************************
'* Nombre           : pImprimir
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPRIME LOS DATOS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pImprimir()
    On Error GoTo error
    Dim oPrint As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    oPrint.Imprimir_x_VSFlexGrid Fg1, "Resumen de Pago a Personal", "", "Del: " & TxtFecha(0).valor & " Al: " & TxtFecha(1).valor, True, True
    Set oPrint = Nothing
    Me.MousePointer = vbDefault
    Exit Sub

error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"
End Sub


Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    On Error GoTo error
    
    Dim mTipoConsulta As Integer
    If SeEjecuto = True Then Exit Sub
    SeEjecuto = True
    
    TxtFecha(0).valor = Date
    TxtFecha(1).valor = Date
    txt_cb(0).Text = ""
    lbl_cb(0).Caption = ""
    
    '*******************************
    txt_cb(1).Text = ""
    lbl_cb(1).Caption = ""
    '*******************************
    txt_cb(2).Text = ""
    lbl_cb(2).Caption = ""
    '*******************************
    TxtFecha(0).SetFocus
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
    SeEjecuto = False
    CentrarFrm Me
    iniciarCampos
    
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
    Exit Sub
    
error:
    SHOW_ERROR
End Sub

Private Sub iniciarCampos()
    Fg1.AllowUserResizing = flexResizeColumns
    Fg1.AutoSearch = flexSearchFromTop
    Fg1.ExplorerBar = flexExSortShow
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.ForeColorSel = &H80000005
    Fg1.BackColorSel = &H80&
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub

    If Me.Height > 3000 Then
        Fg1.Top = 1650
        Fg1.Width = Me.Width - 150
        Fg1.Height = Me.Height - 2050
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BAND_INTERRUMPIR = True
End Sub

'*****************************************************************************************************
'* Nombre           : fValidarConsulta
'* Tipo             : FUNCION
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
    
    ' tipo de planilla
    If chk(0).Value = 0 And chk(1).Value = 0 And chk(4).Value = 0 Then
        chk(0).Value = 1
        chk(1).Value = 1
        chk(4).Value = 1
    End If
    
    ' tipo de horario
    If chk(0).Value = 0 And chk(1).Value = 0 Then
        chk(2).Value = 1
        chk(3).Value = 1
    End If
    fValidarConsulta = True
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

'*****************************************************************************************************
'* Nombre           : pExportarExcel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EXPORTA A MS EXCEL LOS DATOS DEL CONTROL Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pExportarExcel()
    On Error GoTo error
    Dim oExport As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "Resumen de Pago a Personal", "Del: " & TxtFecha(0).valor & " Al: " & TxtFecha(1).valor, , "Resumen de Pago"
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
    
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportarExcel"
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
                
        Case 1: 'Tipo de Planilla
            nTitulo = "Buscando Tipo de Planilla"
            
            nSQL = "SELECT pla_tipoplanilla.id, pla_tipoplanilla.descripcion AS nombre, pla_tipoplanilla.id As cod " _
                + vbCr + "FROM pla_tipoplanilla;"
        Case 2 '--                 personal
            nTitulo = "Buscando Personal"
            nSQL = "SELECT pla_empleados.id, pla_empleados.nombre , pla_empleados.id AS cod " _
                + vbCr + " FROM pla_empleados " _
                + vbCr + " GROUP BY pla_empleados.id, pla_empleados.nombre, pla_empleados.id " _
                + vbCr + " ORDER BY pla_empleados.nombre;"
        
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
    '***************************************
    If KeyAscii = 13 Then
'        If Index <> 1 Then
        SendKeys vbTab
'        Else
'            If Fg1.Rows >= 2 Then
'                Fg1.Row = 1: Fg1.Col = 1
'            Else
'                Fg1.Row = Fg1.Rows - 1: Fg1.Col = 1
'            End If
'            Fg1.SetFocus
'        End If
        Exit Sub
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
    '***************************************
End Sub

Private Sub txt_cb_Validate(Index As Integer, Cancel As Boolean)
    If txt_cb(Index).Text = "" Then Exit Sub
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error
    Select Case Index
        Case 0 '--area
            nSQL = "SELECT mae_area.id, mae_area.descripcion AS nombre, mae_area.id AS cod, mae_area.id AS idarea " _
                + vbCr + " FROM pro_area INNER JOIN mae_area ON pro_area.idarea = mae_area.id " _
                + vbCr + "WHERE (mae_area.id = " & NulosN(txt_cb(Index).Text) & ");"
        
        '****************************************************
        Case 1 ' Tipo de Planilla
            nSQL = "SELECT pla_tipoplanilla.id, pla_tipoplanilla.descripcion AS nombre, pla_tipoplanilla.id As cod " _
                + vbCr + "FROM pla_tipoplanilla " _
                + vbCr + "WHERE (pla_tipoplanilla.id = " & NulosN(txt_cb(Index).Text) & ");"
        '****************************************************
        
        Case Else
            Exit Sub
    End Select

    If xCon.State = 0 Then GoTo SALIR
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo SALIR

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cb(Index).Text = NulosC(RstTmp.Fields(0))   '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = NulosC(RstTmp.Fields(1)) '--NOMBRE
        lbl_cod(Index).Caption = NulosN(RstTmp.Fields(2)) '--CODIGO
        lbl_cb(Index).ToolTipText = NulosC(RstTmp.Fields(1)) '--NOMBRE
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

'*****************************************************************************************************
'* Nombre           : pConsultar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA LA CONSULTA DE LA PLANILLA DE PAGOS DE LOS TRABAJADORES
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pConsultar()
    Dim Rst As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLIdArea As String    '--almacenar el filtro por area
    Dim nSQLIdTipPla As String    ' Filtro para Tipo de Planilla
    Dim nSQLIdEmp As String '--Filtro para personal
    '*********************************************************************
    
    Dim nSQLTipo As String      '--almacenara el filtro por detajo,horas o lineas
    Dim nSQLTurno As String     '--almacenara el filtro por turno
    Dim nSQLCampo As String     '--almacenara los campos para mostrar en reporte
    Dim nSQLCampoGr As String   '--almacenara los campos para aplicar los grupos en reporte
    Dim xColTotal As Integer    '--indica la posicion de la fecha para poner el total al final de los registros
        
    BAND_INTERRUMPIR = False
    If fValidarConsulta() = False Then Exit Sub
    
    If NulosN(lbl_cod(0).Caption) <> 0 Then nSQLIdArea = " and pro_pagos.idarea= " & NulosN(lbl_cod(0).Caption)
    
    '**********************************************
    If NulosN(lbl_cod(1).Caption) <> 0 Then nSQLIdTipPla = " and pla_empleados.idtippla= " & NulosN(lbl_cod(1).Caption)
    '**********************************************
    
    If NulosN(lbl_cod(2).Caption) <> 0 Then nSQLIdEmp = " and pla_empleados.id= " & NulosN(lbl_cod(2).Caption)


'--chk(0) x destajo
'--chk(1) x horas
'--chk(4) x linea

'--chk(2) x dia
'--chk(3) x noche

'--chk(5) x aplica grupo


    If chk(0).Value = 1 And chk(1).Value = 1 And chk(4).Value = 1 Then
    
    ElseIf chk(0).Value = 1 And chk(1).Value = 1 And chk(4).Value = 0 Then '--destajo y horas
        nSQLTipo = " and pro_pagos.tipo in (2,1) "
        
    ElseIf chk(0).Value = 1 And chk(1).Value = 0 And chk(4).Value = 1 Then '--destajo y lineas
        nSQLTipo = " and pro_pagos.tipo in (2,3) "
        
    ElseIf chk(0).Value = 0 And chk(1).Value = 1 And chk(4).Value = 1 Then '--horas y lineas
        nSQLTipo = " and pro_pagos.tipo in (1,3) "
        
    ElseIf chk(0).Value = 1 And chk(1).Value = 0 And chk(4).Value = 0 Then '--destajo
        nSQLTipo = " and pro_pagos.tipo in (2) "
        
    ElseIf chk(0).Value = 0 And chk(1).Value = 1 And chk(4).Value = 0 Then '--horas
        nSQLTipo = " and pro_pagos.tipo in (1) "
        
    ElseIf chk(0).Value = 0 And chk(1).Value = 0 And chk(4).Value = 1 Then '--lineas
        nSQLTipo = " and pro_pagos.tipo in (3) "
    End If
    
    
    If chk(2).Value = 1 And chk(3) = 1 Then
        
    ElseIf chk(2).Value = 1 Then
        nSQLTurno = " and pro_pagos.turno=1 "
    ElseIf chk(3).Value = 1 Then
        nSQLTurno = " and pro_pagos.turno=2 "
        
    End If
    
    Dim nSQLCampos As String
    Dim nSQLGrupo As String
    
    '--agrupando por tipo
    If ChkGrupo.Value = 1 Then
        nSQLCampos = "IIf([pro_pagos].[tipo]=1,'Horas' , iif([pro_pagos].[tipo]=2,'Destajo','Linea')) AS Tipo, "
        nSQLGrupo = "IIf([pro_pagos].[tipo]=1,'Horas' , iif([pro_pagos].[tipo]=2,'Destajo','Linea')) ,"
    End If
    
    '--agrupando por areas
    If chkArea.Value = 1 Then
        nSQLCampos = nSQLCampos & "mae_area.descripcion as area, "
        nSQLGrupo = nSQLGrupo & "mae_area.descripcion, "
    End If

    nSQL = "TRANSFORM Sum(pro_pagos.impbrut) AS SumaDeimpbrut " _
        + vbCr + "SELECT " & nSQLCampos & " pla_tipoplanilla.descripcion AS Planilla, pla_empleados.numdoc, pla_empleados.nombre AS personal, pla_empleados.fching, Sum(pro_pagos.impbrut) AS Total " _
        + vbCr + "FROM (mae_area RIGHT JOIN (pro_pagos INNER JOIN pla_empleados ON pro_pagos.idemp = pla_empleados.id) ON mae_area.id = pro_pagos.idarea) LEFT JOIN pla_tipoplanilla ON pla_empleados.idtippla = pla_tipoplanilla.id " _
        + vbCr + "WHERE (((pro_pagos.fchtra) Between CDate('" & TxtFecha(0).valor & "') And CDate('" & TxtFecha(1).valor & "'))) " & nSQLIdArea & nSQLIdTipPla & nSQLTipo & nSQLTurno & nSQLIdEmp _
        + vbCr + "GROUP BY " & nSQLGrupo & " pla_tipoplanilla.descripcion, pro_pagos.idemp, pla_empleados.numdoc, pla_empleados.nombre, pla_empleados.fching " _
        + vbCr + "ORDER BY pla_empleados.nombre " _
        + vbCr + "PIVOT pro_pagos.fchtra; "
    
    RST_Busq Rst, nSQL, xCon
    If Rst.State = 0 Then Exit Sub
    
    Dim xFila&
    Dim xCol As Integer
    Dim xCantCampo&
    Dim xCampo As String
    
    Fg1.Rows = 1
    Fg1.Cols = 1
    Fg1.RowHeight(0) = 300
    DoEvents
    Do While Not Rst.EOF
        xCol = 0
        Fg1.Rows = Fg1.Rows + 1
        For xCantCampo = 0 To Rst.Fields.Count - 1
            xCol = xCol + 1
            If BAND_INTERRUMPIR = True Then Exit Sub
            
            xCampo = Rst.Fields(xCantCampo).Name
            
            '--colocando el formato del reporte
            If Rst.Bookmark = 1 Then
                Fg1.Cols = Fg1.Cols + 1
                '*********************************************************
                If IsDate(xCampo) Then xCampo = Format(xCampo, "dd/mm/yy")
                '*********************************************************
                Fg1.TextMatrix(0, xCol) = xCampo
                Select Case LCase(xCampo)
                    Case "idemp", "idarea"
                        Fg1.ColWidth(xCol) = 0:
                    Case "tipo"
                        Fg1.ColWidth(xCol) = 700:       Fg1.ColAlignment(xCol) = flexAlignLeftCenter:    Fg1.Row = 0: Fg1.Col = xCol: Fg1.CellAlignment = flexAlignLeftCenter
                    Case "area"
                        Fg1.ColWidth(xCol) = 900:       Fg1.ColAlignment(xCol) = flexAlignLeftCenter:    Fg1.Row = 0: Fg1.Col = xCol: Fg1.CellAlignment = flexAlignLeftCenter
                        Fg1.TextMatrix(0, xCol) = "Area"
                    
                    '******************************************************
                    Case "planilla"
                        Fg1.ColWidth(xCol) = 1500:       Fg1.ColAlignment(xCol) = flexAlignLeftCenter:    Fg1.Row = 0: Fg1.Col = xCol: Fg1.CellAlignment = flexAlignLeftCenter
                        Fg1.TextMatrix(0, xCol) = "T. Planilla"
                    '******************************************************
                    
                    Case "numdoc"
                        Fg1.ColWidth(xCol) = 900:       Fg1.ColAlignment(xCol) = flexAlignCenterCenter:    Fg1.Row = 0: Fg1.Col = xCol: Fg1.CellAlignment = flexAlignLeftCenter
                        Fg1.TextMatrix(0, xCol) = "DNI"
                    Case "personal"
                        Fg1.ColWidth(xCol) = 2400:       Fg1.ColAlignment(xCol) = flexAlignLeftCenter:    Fg1.Row = 0: Fg1.Col = xCol: Fg1.CellAlignment = flexAlignLeftCenter
                        Fg1.TextMatrix(0, xCol) = "Apellidos y Nombres"
                        If chkArea.Value = 0 Then Fg1.ColWidth(xCol) = 3000
                    Case "fching"
                        Fg1.ColWidth(xCol) = 1100:       Fg1.ColAlignment(xCol) = flexAlignCenterCenter:    Fg1.Row = 0: Fg1.Col = xCol: Fg1.CellAlignment = flexAlignLeftCenter
                        Fg1.TextMatrix(0, xCol) = "Fch Ingreso"
                        xColTotal = xCol
                    Case Else
                        Fg1.ColWidth(xCol) = 900:       Fg1.ColAlignment(xCol) = flexAlignRightCenter:  Fg1.Row = 0: Fg1.Col = xCol: Fg1.CellAlignment = flexAlignCenterCenter
                        If chkArea.Value = 0 Then Fg1.ColWidth(xCol) = 1000
                        
                End Select
                '--
                FORMATO_CELDA Fg1, 0, xCol, , True, &HC0C0C0
                
            End If
            
            ' colocando los valores en la grilla
            Select Case LCase(xCampo)
                ' los dias y el total
                Case InStr(xCampo, "/"), "total"
                    If NulosN(Rst.Fields(xCampo)) <> 0 Then
                        Fg1.TextMatrix(Fg1.Rows - 1, xCol) = Format(NulosN(Rst.Fields(xCampo)), FORMAT_MONTO)
                    End If
                 '--otros datos
                Case Else
                    '************************************************************************************
                    Fg1.TextMatrix(Fg1.Rows - 1, xCol) = NulosC(Rst.Fields(Format(xCampo, "dd/mm/yyyy")))
                    '************************************************************************************
            End Select
        Next
        Rst.MoveNext
    Loop
    
    '--colocar los totales por dia
    If Fg1.Rows > 1 Then
        Fg1.Rows = Fg1.Rows + 1
        

        FORMATO_CELDA Fg1, Fg1.Rows - 1, xColTotal, , True, , "Totales"
        For xCol = xColTotal + 1 To Fg1.Cols - 1
            FORMATO_CELDA Fg1, Fg1.Rows - 1, xCol, , True, , Format(GRID_SUMAR_COL(Fg1, xCol), FORMAT_MONTO)
        Next

    End If
End Sub
