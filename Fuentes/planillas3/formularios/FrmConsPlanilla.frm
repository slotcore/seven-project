VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsPlanilla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilas - Consulta"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "( Periodo )"
      Height          =   585
      Left            =   9870
      TabIndex        =   16
      Top             =   345
      Width           =   1980
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
         Left            =   120
         TabIndex        =   17
         Top             =   255
         Width           =   1740
      End
   End
   Begin VB.Frame Frame1 
      Height          =   585
      Left            =   30
      TabIndex        =   8
      Top             =   345
      Width           =   9795
      Begin VB.CommandButton cb 
         Height          =   225
         Index           =   1
         Left            =   1185
         Picture         =   "FrmConsPlanilla.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Seleccione el Personal (Opcional)"
         Top             =   210
         Width           =   210
      End
      Begin VB.CommandButton cb 
         Height          =   225
         Index           =   0
         Left            =   5460
         Picture         =   "FrmConsPlanilla.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Seleccione el Personal (Opcional)"
         Top             =   210
         Width           =   210
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   0
         Left            =   4920
         MaxLength       =   20
         TabIndex        =   10
         Text            =   "txt_cb(0)"
         Top             =   180
         Width           =   780
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   1
         Left            =   645
         MaxLength       =   20
         TabIndex        =   19
         Text            =   "txt_cb(1)"
         Top             =   180
         Width           =   780
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
         Height          =   285
         Index           =   1
         Left            =   2820
         TabIndex        =   21
         Top             =   150
         Visible         =   0   'False
         Width           =   975
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
         Height          =   285
         Index           =   1
         Left            =   1425
         TabIndex        =   22
         Top             =   180
         Width           =   2295
      End
      Begin VB.Label lbl_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   20
         Top             =   240
         Width           =   420
      End
      Begin VB.Label lbl_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Personal"
         Height          =   195
         Index           =   0
         Left            =   4215
         TabIndex        =   12
         Top             =   240
         Width           =   615
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
         Height          =   285
         Index           =   0
         Left            =   9240
         TabIndex        =   11
         Top             =   150
         Visible         =   0   'False
         Width           =   975
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
         Height          =   285
         Index           =   0
         Left            =   5700
         TabIndex        =   13
         Top             =   180
         Width           =   2955
      End
   End
   Begin VB.Frame fra_barra 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   795
      Left            =   2775
      TabIndex        =   1
      Top             =   3195
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar barra 
         Height          =   285
         Left            =   105
         TabIndex        =   2
         Top             =   330
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblbarra 
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
         Index           =   1
         Left            =   4275
         TabIndex        =   4
         Top             =   90
         Width           =   1530
      End
      Begin VB.Label lblbarra 
         AutoSize        =   -1  'True
         Caption         =   "Procesando Planillas"
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
         Left            =   165
         TabIndex        =   3
         Top             =   90
         Width           =   1725
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   960
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   5925
         X2              =   5925
         Y1              =   -15
         Y2              =   915
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   -30
         X2              =   5940
         Y1              =   15
         Y2              =   30
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   5910
         Y1              =   780
         Y2              =   765
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
         Left            =   4830
         Top             =   135
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
               Picture         =   "FrmConsPlanilla.frx":0264
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsPlanilla.frx":07A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsPlanilla.frx":0B3A
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsPlanilla.frx":0CBE
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsPlanilla.frx":1112
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsPlanilla.frx":122A
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsPlanilla.frx":176E
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsPlanilla.frx":1CB2
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsPlanilla.frx":1DC6
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsPlanilla.frx":1EDA
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsPlanilla.frx":232E
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsPlanilla.frx":249A
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsPlanilla.frx":29E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsPlanilla.frx":2CFC
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6585
      Left            =   -15
      TabIndex        =   5
      Top             =   945
      Width           =   11910
      _cx             =   21008
      _cy             =   11615
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
      FrontTabForeColor=   -2147483630
      Caption         =   "      Detalle    |     Resumen     "
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   6165
         Left            =   45
         TabIndex        =   7
         Top             =   45
         Width           =   11820
         Begin VSFlex7Ctl.VSFlexGrid Fg 
            Height          =   5985
            Index           =   0
            Left            =   105
            TabIndex        =   14
            Top             =   90
            Width           =   11595
            _cx             =   20452
            _cy             =   10557
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
            Rows            =   1
            Cols            =   1
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmConsPlanilla.frx":308E
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
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   6165
         Left            =   12555
         TabIndex        =   6
         Top             =   45
         Width           =   11820
         Begin VSFlex7Ctl.VSFlexGrid Fg 
            Height          =   5985
            Index           =   1
            Left            =   105
            TabIndex        =   15
            Top             =   90
            Width           =   11610
            _cx             =   20479
            _cy             =   10557
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
            Rows            =   1
            Cols            =   1
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmConsPlanilla.frx":3173
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
End
Attribute VB_Name = "FrmConsPlanilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean
Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE
Dim nSQLPivot As String

Dim mMesActivo As Integer '--indica el mes activo

Private Sub CambiarMes()
    mMesActivo = SeleccionaMes(xCon)
    lblperiodo.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
End Sub

Private Sub pExportar()
    If AnoTra = "" Then
        MsgBox "No Hay Año de trabajo" + vbCr + "No puede continuar", vbExclamation, xTitulo
        Exit Sub
    End If

    If mMesActivo = 0 Then
        MsgBox "Seleccione el Periodo de Consulta", vbExclamation, xTitulo
        Exit Sub
    End If
    On Error GoTo error
    
    Dim oExport As New SGI2_funciones.formularios
    Dim mIndex As Integer
    Dim nTitulo As String
    Dim nPeriodo As String
    Dim nTitulo1 As String
    If TabOne1.CurrTab = 0 Then
            mIndex = 0
            nTitulo = "Consulta de Planillas"
    Else
            mIndex = 1
            nTitulo = "Consulta de Planillas - Resumen"
    End If
    
    nPeriodo = "Periodo: " + lblperiodo.Caption
    If NulosN(lbl_cod(0).Caption) <> 0 Then
         nTitulo1 = "Personal: " & StrConv(lbl_cb(0).Caption, 3)
    End If

    Me.MousePointer = vbHourglass
    oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg(mIndex), nTitulo, nPeriodo, nTitulo1, "Consulta de Planillas"
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Exportar"
End Sub


Private Sub pImprimir()
    If AnoTra = "" Then
        MsgBox "No Hay Año de trabajo" + vbCr + "No puede continuar", vbExclamation, xTitulo
        Exit Sub
    End If

    If mMesActivo = 0 Then
        MsgBox "Seleccione el Periodo de Consulta", vbExclamation, xTitulo
        Exit Sub
    End If
    
    On Error GoTo error

    Dim oPrint  As New SGI2_funciones.formularios
    Dim mIndex As Integer
    Dim nTitulo As String
    Dim nPeriodo As String
    Dim nTitulo1 As String
    If TabOne1.CurrTab = 0 Then
            mIndex = 0
            nTitulo = "REPORTE DE PLANILLAS"
            
    Else
            mIndex = 1
            nTitulo = "RESUMEN DE PLANILLAS"
    End If
    
    If txt_cb(1).Text <> 0 Then nTitulo = nTitulo & " - " & UCase(lbl_cb(1).Caption)
    
    nPeriodo = "Periodo: " + lblperiodo.Caption
    
    If NulosN(lbl_cod(0).Caption) <> 0 Then
        nTitulo1 = "Personal: " & StrConv(lbl_cb(0).Caption, 3)
    End If
    
    Me.MousePointer = vbHourglass
    oPrint.Imprimir_x_VSFlexGrid Fg(mIndex), nTitulo, nTitulo1, nPeriodo, True, True
    Set oPrint = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Set oPrint = Nothing
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Exportar"
End Sub

Private Sub pConsultar()
    ''''''''''''
    If AnoTra = "" Then
        MsgBox "No Hay Año de trabajo" + vbCr + "No puede continuar", vbExclamation, xTitulo
        Exit Sub
    End If

    If mMesActivo = 0 Then
        MsgBox "Seleccione el Periodo de Consulta", vbExclamation, xTitulo
        Exit Sub
    End If

    '''''''''''
    BAND_INTERRUMPIR = False
    
    '----
    fra_barra.Visible = True
    fra_barra.Top = 3195
    fra_barra.Left = 2775
    '----
    Me.TabOne1.CurrTab = 0
    BAND_INTERRUMPIR = False
    pCargarDetalle
    '--SI SE NTERRUMPE EL PROCESO => SALIR
    If BAND_INTERRUMPIR = True Then GoTo salir:
    '-----------------------------------------------
salir:
    fra_barra.Visible = False
    If BAND_INTERRUMPIR = True Then
        MsgBox "La consulta fue interrumpida", vbInformation, xTitulo
    End If
        
End Sub


Private Sub Form_Activate()
    If SeEjecuto = False Then
        mMesActivo = xMes
        
        SeEjecuto = True
        txt_cb(0).Text = ""
        txt_cb(1).Text = ""
        
        LimpiaText lblperiodo
        lblperiodo.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
        
        TabOne1.CurrTab = 0
        
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    End If
End Sub

Private Sub Form_Load()
    CentrarFrm Me
    SeEjecuto = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 3 Then CambiarMes
    If Button.Index = 5 Then pExportar
    If Button.Index = 6 Then pImprimir
    If Button.Index = 8 Then
        Unload Me
    End If
End Sub

'****************************************************************************************
Private Sub cb_Click(Index As Integer)
    On Error GoTo error
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim xCampos() As String

    
    Select Case Index
    Case 0
        pBuscarPersonal xRs, False
    Case 1
        ReDim xCampos(2, 3) As String
        xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombres":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
        xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":         xCampos(1, 2) = "400":    xCampos(1, 3) = "N"
        
        nSQL = "SELECT mae_cargo.id, mae_cargo.descripcion AS nombres, mae_cargo.id AS cod " _
            + vbCr + " FROM mae_cargo;"
        RST_Busq xRs, nSQL, xCon
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "", "nombres", "nombres", Principio

    End Select
    
    
    
    
    If xRs.State = 1 Then
        txt_cb(Index) = xRs.Fields("id") & "" '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = xRs.Fields("nombres") & "" '--NOMBRE
        lbl_cod(Index).Caption = xRs.Fields("id") & "" '--CODIGO
        lbl_cb(Index).ToolTipText = xRs.Fields("nombres") & "" '--NOMBRE
        txt_cb(Index).SetFocus
    End If
    Set xRs = Nothing
Exit Sub
error:
    Set xRs = Nothing
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
        SendKeys vbTab
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
        Case 0 '--TIPO DE TRABAJADOR
            nSQL = "SELECT pla_empleados.id, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, pla_empleados.id as cod, pla_empleados.numdoc, mae_dociden.abrev AS tipodoc, mae_sexo.abrev AS sexo, Format([pla_empleados].[fchnac],'dd/mm/yyyy') AS fchnac, pla_empleados.numtel, pla_empleados.email " _
                + vbCr + " FROM mae_sexo RIGHT JOIN (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) ON mae_sexo.id = pla_empleados.idsex " _
                + vbCr + " WHERE  pla_empleados.id  = " & NulosC(txt_cb(Index).Text) & ";"
        Case 1 '--cargo
        
        
            nSQL = "SELECT mae_cargo.id, mae_cargo.descripcion AS nombre, mae_cargo.id AS cod " _
                + vbCr + " FROM mae_cargo " _
                + vbCr + " WHERE mae_cargo.id = " & NulosN(txt_cb(Index).Text) & ";"
        
    End Select

    If xCon.State = 0 Then GoTo salir
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cb(Index).Text = RstTmp.Fields(0) & ""  '--TEXTO A MOSTRAR
        lbl_cb(Index).Caption = RstTmp.Fields(1) & "" '--NOMBRE
        lbl_cod(Index).Caption = RstTmp.Fields(2) & "" '--CODIGO
        lbl_cb(Index).ToolTipText = RstTmp.Fields(1) & "" '--NOMBRE
    Else
        txt_cb(Index).Text = "":    lbl_cb(Index).Caption = "":    lbl_cod(Index).Caption = ""
    End If
    '--------------
    Set RstTmp = Nothing
    Exit Sub
error:
    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "txt_cb_Validate (" + CStr(Index) + ")"
    Exit Sub
salir:
    Set RstTmp = Nothing
    txt_cb(Index).Text = ""
End Sub

'****************************************************************************************
Private Sub pCargarDetalle()
    
    Dim nSQL As String
    Dim RstTmp As New ADODB.Recordset
    Dim nSQLIdEmp As String
    
    If NulosN(lbl_cod(0).Caption) <> 0 Then
        nSQLIdEmp = " and pla_boleta.idemp = " & NulosN(lbl_cod(0).Caption)
    End If
    If NulosN(lbl_cod(1).Caption) <> 0 Then
        nSQLIdEmp = nSQLIdEmp & " and pla_empleados.idcargo = " & NulosN(lbl_cod(1).Caption)
    End If
    
    
    '----
    lblbarra(0).Caption = "Procesando Detalle por Planilla"
    Me.barra.Max = 10
    Me.barra.Min = 1
    Me.barra.Value = 1
    '--limpiar la grilla
    Fg(0).SelectionMode = flexSelectionByRow
    Fg(0).Rows = Fg(0).FixedRows
    DoEvents
    '*************************************************************************************
    '--Generar Consulta datos del personal
    
    nSQL = "SELECT pla_boleta.idemp,mae_categoria.nomcor AS Cat, mae_dociden.abrev AS TipDoc, pla_empleados.numdoc as NumDoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS Nombres, mae_cargo.descripcion AS Cargo " _
        + vbCr + " FROM (mae_dociden RIGHT JOIN (mae_cargo RIGHT JOIN pla_empleados ON mae_cargo.id = pla_empleados.idcargo) ON mae_dociden.id = pla_empleados.idtipdoc) INNER JOIN (mae_moneda RIGHT JOIN (pla_boleta LEFT JOIN mae_categoria ON pla_boleta.idcat = mae_categoria.id) ON mae_moneda.id = pla_boleta.idmon) ON pla_empleados.id = pla_boleta.idemp " _
        + vbCr + " WHERE (((pla_boleta.ano)=" & AnoTra & ") AND ((pla_boleta.idmes)=" & mMesActivo & ")) " & nSQLIdEmp & " ;"
    
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.RecordCount = 0 Then
        MsgBox "No hay Registros", vbInformation, xTitulo
        Exit Sub
    End If
    Fg(0).Cols = 1
    Fg(0).Rows = 2
    Fg(0).FixedRows = 2
    '--poner encabezado
    Dim mColRst As Integer, mColGrid As Integer, mRowGrid As Integer
    Dim mColTmp As Integer         '--almacena la ultima columna despues de datos,ingresos,descuentos,aportes
    Dim mColIng As Integer, mColDesc As Integer, mColApo As Integer      '--indican las columnas del total
    Dim mColInicio As Integer      '--almacenar la posicion inicial donde empieza a colocar los acumulados
    With Fg(0)
        .FrozenCols = 0
        Me.barra.Value = 1
        For mColRst = 0 To RstTmp.Fields.Count - 1
            .Cols = .Cols + 1
            mColGrid = .Cols - 1
            .TextMatrix(1, mColGrid) = RstTmp.Fields(mColRst).Name:
            Select Case LCase(RstTmp.Fields(mColRst).Name)
                Case "nombres"
                    .ColWidth(mColGrid) = 2500:  .ColAlignment(mColGrid) = flexAlignLeftCenter: .Row = 1: .Col = mColGrid:  .CellAlignment = flexAlignLeftCenter
                Case "cat"
                    .ColWidth(mColGrid) = 0:  .ColAlignment(mColGrid) = flexAlignLeftCenter:  .Row = 1: .Col = mColGrid:  .CellAlignment = flexAlignLeftCenter
                Case "cargo"
                    .ColWidth(mColGrid) = 900:  .ColAlignment(mColGrid) = flexAlignLeftCenter:  .Row = 1: .Col = mColGrid:  .CellAlignment = flexAlignLeftCenter
                    If NulosN(txt_cb(1).Text) <> 0 Then .ColWidth(mColGrid) = 0
                    
                Case "tipdoc"
                    .ColWidth(mColGrid) = 600:  .ColAlignment(mColGrid) = flexAlignCenterCenter: .Row = 1: .Col = mColGrid: .CellAlignment = flexAlignCenterCenter
                Case "numdoc"
                    .ColWidth(mColGrid) = 800:  .ColAlignment(mColGrid) = flexAlignCenterCenter: .Row = 1: .Col = mColGrid: .CellAlignment = flexAlignCenterCenter
                Case "idemp"
                    .ColWidth(mColGrid) = 0:  .ColAlignment(mColGrid) = flexAlignCenterCenter:  .Row = 1: .Col = mColGrid:  .CellAlignment = flexAlignCenterCenter
                Case Else
                    .ColWidth(mColGrid) = 800:  .ColAlignment(mColGrid) = flexAlignLeftCenter:  .Row = 1: .Col = mColGrid:  .CellAlignment = flexAlignLeftCenter
            End Select
            
        Next mColRst
        .FrozenCols = mColGrid
        GRID_COMBINAR Fg(0), 0, 1, 0, .Cols - 1, "Datos del Personal", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC, True
    End With
    '--cargar los datos
    Me.barra.Value = 2
    RstTmp.MoveFirst
    Do While Not RstTmp.EOF
        DoEvents
        Fg(0).Rows = Fg(0).Rows + 1
        mRowGrid = Fg(0).Rows - 1
        mColGrid = 1
        For mColRst = 0 To RstTmp.Fields.Count - 1
            Select Case LCase(RstTmp.Fields(mColRst).Name)
                Case "nombres", "cargo", "categoria", "numdoc", "tipdoc"
                    Fg(0).TextMatrix(mRowGrid, mColGrid) = NulosC(RstTmp.Fields(mColRst))
                Case Else
                    Fg(0).TextMatrix(mRowGrid, mColGrid) = NulosC(RstTmp.Fields(mColRst))
            End Select
            mColGrid = mColGrid + 1
        Next mColRst
        RstTmp.MoveNext
    Loop
    
    Set RstTmp = Nothing
    
    mColInicio = Fg(0).Cols - 1
    
    '*************************************************************************************
    
    nSQLIdEmp = Replace(nSQLIdEmp, "pla_empleados", "pla_boleta")
    
    '---Generar Consulta de los Ingresos
    nSQL = "TRANSFORM Sum(pla_boletadet.imptot) AS total " _
        + vbCr + " SELECT pla_boleta.idemp, Sum(pla_boletadet.imptot) AS acumulado " _
        + vbCr + " FROM (pla_conceptotipo INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo) INNER JOIN ((mae_moneda RIGHT JOIN pla_boleta ON mae_moneda.id = pla_boleta.idmon) INNER JOIN pla_boletadet ON pla_boleta.id = pla_boletadet.idbol) ON pla_concepto.id = pla_boletadet.idcpto " _
        + vbCr + " WHERE (((pla_boleta.ano)=" & AnoTra & ") AND ((pla_boleta.idmes)=" & mMesActivo & ") AND ((pla_conceptotipo.idcat)=1) AND ((pla_concepto.aplanilla)=-1)) " & nSQLIdEmp & " " _
        + vbCr + " GROUP BY pla_boleta.idemp " _
        + vbCr + " ORDER BY pla_boleta.idemp " _
        + vbCr + " PIVOT pla_concepto.nomcorto; "

    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.RecordCount <> 0 Then
        pCargarDetalleDet RstTmp, "Ingresos", 3
        mColIng = Fg(0).Cols - 1
    End If
    Set RstTmp = Nothing
    If BAND_INTERRUMPIR = True Then Exit Sub '--si se interrumpe
    '*************************************************************************************
    
    '---Generar Consulta de los Descuentos
    nSQL = "TRANSFORM Sum(pla_boletadet.imptot) AS total " _
        + vbCr + " SELECT pla_boleta.idemp, Sum(pla_boletadet.imptot) AS acumulado " _
        + vbCr + " FROM (pla_conceptotipo INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo) INNER JOIN ((mae_moneda RIGHT JOIN pla_boleta ON mae_moneda.id = pla_boleta.idmon) INNER JOIN pla_boletadet ON pla_boleta.id = pla_boletadet.idbol) ON pla_concepto.id = pla_boletadet.idcpto " _
        + vbCr + " WHERE (((pla_boleta.ano)=" & AnoTra & ") AND ((pla_boleta.idmes)=" & mMesActivo & ") AND ((pla_conceptotipo.idcat)=3) AND ((pla_concepto.aplanilla)=-1)) OR (((pla_boleta.ano)=" & AnoTra & ") AND ((pla_boleta.idmes)=" & mMesActivo & ") AND ((pla_conceptotipo.idcat)=2) AND ((pla_concepto.idtipo)=9)) " & nSQLIdEmp & " " _
        + vbCr + " GROUP BY pla_boleta.idemp " _
        + vbCr + " PIVOT pla_concepto.nomcorto; "
        
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.RecordCount <> 0 Then
        pCargarDetalleDet RstTmp, "Descuentos", 5
        mColDesc = Fg(0).Cols - 1
    End If
    Set RstTmp = Nothing
    If BAND_INTERRUMPIR = True Then Exit Sub '--si se interrumpe
    '*************************************************************************************
    
    '---Generar Consulta de los Aportes
    nSQL = "TRANSFORM Sum(pla_boletadet.imptot) AS total " _
        + vbCr + " SELECT pla_boleta.idemp, Sum(pla_boletadet.imptot) AS acumulado " _
        + vbCr + " FROM (pla_conceptotipo INNER JOIN pla_concepto ON pla_conceptotipo.id = pla_concepto.idtipo) INNER JOIN ((mae_moneda RIGHT JOIN pla_boleta ON mae_moneda.id = pla_boleta.idmon) INNER JOIN pla_boletadet ON pla_boleta.id = pla_boletadet.idbol) ON pla_concepto.id = pla_boletadet.idcpto " _
        + vbCr + " WHERE (((pla_boleta.ano)=" & AnoTra & ") AND ((pla_boleta.idmes)=" & mMesActivo & ") AND ((pla_conceptotipo.idcat)=2) AND ((pla_concepto.idtipo)=10) AND ((pla_concepto.aplanilla)=-1)) " & nSQLIdEmp & " " _
        + vbCr + " GROUP BY pla_boleta.idemp, pla_conceptotipo.idcat, pla_concepto.idtipo " _
        + vbCr + " ORDER BY pla_boleta.idemp " _
        + vbCr + " PIVOT pla_concepto.nomcorto; "
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.RecordCount <> 0 Then
        pCargarDetalleDet RstTmp, "Aportes", 7
        mColApo = Fg(0).Cols - 1
    End If
    Set RstTmp = Nothing
    If BAND_INTERRUMPIR = True Then Exit Sub '--si se interrumpe
    '*************************************************************************************
    '--calcular el neto a pagar, totales por columna
    With Fg(0)
        Me.barra.Value = 9
        Fg(0).Cols = Fg(0).Cols + 1
        mColGrid = .Cols - 1
        .TextMatrix(0, mColGrid) = "Neto a"
        .TextMatrix(1, mColGrid) = "Pagar"
        .ColWidth(mColGrid) = 1000:  .ColAlignment(mColGrid) = flexAlignRightCenter: .Row = 0: .Col = mColGrid:  .CellAlignment = flexAlignRightCenter: .CellFontBold = True
        .ColWidth(mColGrid) = 1000:  .ColAlignment(mColGrid) = flexAlignRightCenter: .Row = 1: .Col = mColGrid:  .CellAlignment = flexAlignRightCenter: .CellFontBold = True
        '--calcular el neto a pagar
        For mRowGrid = .FixedRows To .Rows - 1
            .TextMatrix(mRowGrid, .Cols - 1) = Format(NulosN(.TextMatrix(mRowGrid, mColIng)) - NulosN(.TextMatrix(mRowGrid, mColDesc)), FORMAT_MONTO)
        Next
        '--calcular los totales por columna
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, mColInicio) = "Totales"
        For mColGrid = mColInicio + 1 To .Cols - 1
            .TextMatrix(.Rows - 1, mColGrid) = Format(GRID_SUMAR_COL(Fg(0), mColGrid), FORMAT_MONTO)
        Next
        Me.barra.Value = 10
        If mColIng > 0 Then
            '--aplicando el fonfo
            GRID_COLOR_FONDO Fg(0), .FixedRows, mColInicio + 1, .Rows - 1, mColIng - 1, &HE7FEFC  '--concepto ing
            GRID_COLOR_FONDO Fg(0), .FixedRows, mColIng, .Rows - 1, mColIng, &HC8FDF9 '--total ing
            GRID_COLOR_FONDO Fg(0), .FixedRows, mColIng + 1, .Rows - 1, mColDesc - 1, &HC0E0FF '--concepto desc
            GRID_COLOR_FONDO Fg(0), .FixedRows, mColDesc, .Rows - 1, mColDesc, &HB0D8FF  '--total desc
            GRID_COLOR_FONDO Fg(0), .FixedRows, mColDesc + 1, .Rows - 1, mColApo - 1, &HC0C0FF '--concepto apo
            GRID_COLOR_FONDO Fg(0), .FixedRows, mColApo, .Rows - 1, mColApo, &HB0B0FF '--total apo
            GRID_COLOR_FONDO Fg(0), .FixedRows, .Cols - 1, .Rows - 1, .Cols - 1, &HFFD3A8 '--neto a pagar
        End If
        
    End With
    
    
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    '%%% mostrar el Resumen
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    TabOne1.CurrTab = 1
    lblbarra(0).Caption = "Procesando Remunen de Planilla"
    barra.Min = 1
    barra.Max = Fg(0).Rows - 1
    barra.Value = 1
    
    With Fg(1)
        DoEvents
        .SelectionMode = flexSelectionByRow
        .Rows = 2
        .FixedRows = 2
        .FrozenCols = 0
        .Cols = mColInicio + 1 + 4 '(col 0)+(ing + desc + apo + neto a pagar)
        For mColGrid = 1 To mColInicio
            DoEvents
            .TextMatrix(1, mColGrid) = Fg(0).TextMatrix(1, mColGrid)
            .ColWidth(mColGrid) = Fg(0).ColWidth(mColGrid):
            .ColAlignment(mColGrid) = Fg(0).ColAlignment(mColGrid):
            .Row = 1: .Col = mColGrid:
            .CellAlignment = Fg(0).CellAlignment
        Next mColGrid
        
        .TextMatrix(1, .Cols - 4) = "Ingresos":   .ColWidth(.Cols - 4) = 1100: .ColAlignment(.Cols - 4) = flexAlignRightCenter: .Row = 1: .Col = .Cols - 4: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, .Cols - 3) = "Descuentos":  .ColWidth(.Cols - 3) = 1100: .ColAlignment(.Cols - 3) = flexAlignRightCenter: .Row = 1: .Col = .Cols - 3: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, .Cols - 2) = "Aportes":    .ColWidth(.Cols - 2) = 1100: .ColAlignment(.Cols - 2) = flexAlignRightCenter: .Row = 1: .Col = .Cols - 2: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, .Cols - 1) = "Total":      .ColWidth(.Cols - 1) = 1200: .ColAlignment(.Cols - 1) = flexAlignRightCenter: .Row = 1: .Col = .Cols - 1: .CellAlignment = flexAlignRightCenter
        
        .FrozenCols = mColInicio
        GRID_COMBINAR Fg(1), 0, 1, 0, mColInicio, "Datos del Personal", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC, True
        GRID_COMBINAR Fg(1), 0, mColInicio + 1, 0, .Cols - 2, "Totales", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC, True
        GRID_COMBINAR Fg(1), 0, .Cols - 1, 0, .Cols - 1, "Neto a Pagar", flexAlignRightCenter, True, , vbBlack, &HD8E9EC, True
        '--cargar los datos
        For mRowGrid = Fg(0).FixedRows To Fg(0).Rows - 1
            DoEvents
            If BAND_INTERRUMPIR = True Then Exit Sub '--si se interrumpe
            barra.Value = mRowGrid
            .Rows = .Rows + 1
            For mColGrid = 1 To mColInicio
                DoEvents
                If BAND_INTERRUMPIR = True Then Exit Sub '--si se interrumpe
                .TextMatrix(mRowGrid, mColGrid) = Fg(0).TextMatrix(mRowGrid, mColGrid)
            Next mColGrid
            .TextMatrix(mRowGrid, .Cols - 4) = Fg(0).TextMatrix(mRowGrid, mColIng)
            .TextMatrix(mRowGrid, .Cols - 3) = Fg(0).TextMatrix(mRowGrid, mColDesc)
            .TextMatrix(mRowGrid, .Cols - 2) = Fg(0).TextMatrix(mRowGrid, mColApo)
            .TextMatrix(mRowGrid, .Cols - 1) = Fg(0).TextMatrix(mRowGrid, Fg(0).Cols - 1)
        Next mRowGrid
        
        '--aplicando el fonfo
        GRID_COLOR_FONDO Fg(1), .FixedRows, .Cols - 4, .Rows - 1, .Cols - 4, &HC8FDF9 '--total ing
        GRID_COLOR_FONDO Fg(1), .FixedRows, .Cols - 3, .Rows - 1, .Cols - 3, &HB0D8FF  '--total desc
        GRID_COLOR_FONDO Fg(1), .FixedRows, .Cols - 2, .Rows - 1, .Cols - 2, &HB0B0FF '--total apo
        GRID_COLOR_FONDO Fg(1), .FixedRows, .Cols - 1, .Rows - 1, .Cols - 1, &HFFD3A8  '--neto a pagar
        
    End With
    

End Sub

Private Sub pCargarDetalleDet(RstTmp As ADODB.Recordset, nEncabezado As String, barra As Integer)
    Dim mColRst&, mColGrid&, mRowGrid&
    Dim mColTmp As Integer  '--almacena la ultima columna despues de... ,ingresos,descuentos,aportes
    '-----------------------------------------------------------------------------
    '--poner encabezado
    With Fg(0)
        mColTmp = .Cols - 1
        .Cols = .Cols + RstTmp.Fields.Count - 1 '--no considerar el campo de idemp
        Me.barra.Value = barra
        For mColRst = 1 To RstTmp.Fields.Count - 1 '--no considerar la primera columna
            DoEvents
            If BAND_INTERRUMPIR = True Then Exit Sub '--si se interrumpe
            Select Case LCase(RstTmp.Fields(mColRst).Name)
                Case "acumulado"
                    .TextMatrix(1, .Cols - 1) = "Total":
                    .ColWidth(.Cols - 1) = 900:  .ColAlignment(.Cols - 1) = flexAlignRightCenter:
                    .Row = 1: .Col = .Cols - 1:   .CellAlignment = flexAlignRightCenter
                Case Else
                
                    mColGrid = mColTmp + mColRst - 1 '--posicionando en la ultima columna
            
                    .TextMatrix(1, mColGrid) = RstTmp.Fields(mColRst).Name:
                    .ColWidth(mColGrid) = 1000:  .ColAlignment(mColGrid) = flexAlignRightCenter:
                    .Row = 1: .Col = mColGrid:  .CellAlignment = flexAlignRightCenter
            End Select
            
        Next mColRst
        GRID_COMBINAR Fg(0), 0, mColTmp + 1, 0, .Cols - 1, nEncabezado, flexAlignCenterCenter, True, , vbBlack, &HD8E9EC, True
    
        '-----------------------------------------------------------------------------
        '--cargar los datos
        For mRowGrid = .FixedRows To .Rows - 1
            DoEvents
            If BAND_INTERRUMPIR = True Then Exit Sub '--si se interrumpe
            RstTmp.Filter = "idemp=" & NulosN(.TextMatrix(mRowGrid, 1))
            Me.barra.Value = barra + 1
            If RstTmp.RecordCount <> 0 Then
                For mColRst = 1 To RstTmp.Fields.Count - 1 '--no considerar la primera columna
                    Select Case LCase(RstTmp.Fields(mColRst).Name)
                        Case "acumulado"
                            .TextMatrix(mRowGrid, .Cols - 1) = Format(NulosN(RstTmp.Fields(mColRst)), FORMAT_MONTO)
                        Case Else
                            mColGrid = mColTmp + mColRst - 1 '--posicionando en la ultima columna
                            .TextMatrix(mRowGrid, mColGrid) = Format(NulosN(RstTmp.Fields(mColRst)), FORMAT_MONTO)
                    End Select
                Next mColRst
                
            End If
        Next
    End With
    
    
    
End Sub



