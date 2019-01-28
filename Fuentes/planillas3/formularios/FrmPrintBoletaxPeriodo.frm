VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPrintBoletaxPeriodo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilas - Impresión de Boletas de Pagos"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   30
      TabIndex        =   10
      Top             =   345
      Width           =   11790
      Begin VB.Frame Frame4 
         Caption         =   "( Periodo )"
         Height          =   780
         Left            =   9690
         TabIndex        =   28
         Top             =   180
         Width           =   2010
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
            TabIndex        =   29
            Top             =   330
            Width           =   1740
         End
      End
      Begin VB.CommandButton cb 
         Enabled         =   0   'False
         Height          =   225
         Index           =   3
         Left            =   1380
         Picture         =   "FrmPrintBoletaxPeriodo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Seleccione el Personal"
         Top             =   765
         Width           =   210
      End
      Begin VB.TextBox txt_cb 
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   825
         MaxLength       =   20
         TabIndex        =   3
         Text            =   "txt_cb(3)"
         Top             =   735
         Width           =   780
      End
      Begin VB.CommandButton cbA 
         Height          =   225
         Index           =   1
         Left            =   6210
         Picture         =   "FrmPrintBoletaxPeriodo.frx":0132
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   165
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cb 
         Height          =   225
         Index           =   0
         Left            =   1380
         Picture         =   "FrmPrintBoletaxPeriodo.frx":0264
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   165
         Width           =   195
      End
      Begin VB.CommandButton cb 
         Height          =   225
         Index           =   2
         Left            =   1380
         Picture         =   "FrmPrintBoletaxPeriodo.frx":0396
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   465
         Width           =   195
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   2
         Left            =   825
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   2
         Text            =   "txt_cb(1)"
         Top             =   435
         Width           =   780
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   0
         Left            =   825
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   0
         Text            =   "txt_cb(0)"
         Top             =   135
         Width           =   780
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   1
         Left            =   5670
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   1
         Text            =   "txt_cb(4)"
         Top             =   135
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Personal"
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   75
         TabIndex        =   13
         Top             =   855
         Width           =   615
      End
      Begin VB.Label lbl_cod 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cod(3)"
         Enabled         =   0   'False
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
         Index           =   3
         Left            =   4995
         TabIndex        =   12
         Top             =   750
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proceso"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   24
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lbl_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoría"
         Height          =   195
         Index           =   2
         Left            =   75
         TabIndex        =   23
         Top             =   555
         Width           =   705
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
         Height          =   285
         Index           =   2
         Left            =   3495
         TabIndex        =   22
         Top             =   435
         Visible         =   0   'False
         Width           =   1110
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
         Left            =   3495
         TabIndex        =   21
         Top             =   135
         Visible         =   0   'False
         Width           =   1110
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
         Left            =   7395
         TabIndex        =   20
         Top             =   165
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label lbl_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número"
         Height          =   195
         Index           =   1
         Left            =   4935
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   555
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
         Left            =   6420
         TabIndex        =   27
         Top             =   135
         Visible         =   0   'False
         Width           =   2970
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
         Height          =   285
         Index           =   2
         Left            =   1605
         TabIndex        =   26
         Top             =   435
         Width           =   2970
      End
      Begin VB.Label lbl_cb 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cb(3)"
         Enabled         =   0   'False
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
         Index           =   3
         Left            =   1605
         TabIndex        =   14
         Top             =   735
         Width           =   7815
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
         Left            =   1605
         TabIndex        =   25
         Top             =   135
         Width           =   2970
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5700
      Left            =   30
      TabIndex        =   15
      Top             =   1425
      Width           =   11775
      Begin VSFlex7Ctl.VSFlexGrid Fg 
         Height          =   5505
         Index           =   0
         Left            =   75
         TabIndex        =   4
         Top             =   90
         Width           =   11640
         _cx             =   20532
         _cy             =   9710
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
         FormatString    =   $"FrmPrintBoletaxPeriodo.frx":04C8
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
   Begin VB.Frame fra_barra 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   795
      Left            =   2775
      TabIndex        =   6
      Top             =   3195
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar barra 
         Height          =   285
         Left            =   105
         TabIndex        =   7
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
         TabIndex        =   9
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
         TabIndex        =   8
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
      Height          =   345
      Left            =   0
      TabIndex        =   5
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Boletas"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Resumen"
               EndProperty
            EndProperty
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
         Left            =   5040
         Top             =   195
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
               Picture         =   "FrmPrintBoletaxPeriodo.frx":05AD
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrintBoletaxPeriodo.frx":0AF1
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrintBoletaxPeriodo.frx":0E83
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrintBoletaxPeriodo.frx":1007
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrintBoletaxPeriodo.frx":145B
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrintBoletaxPeriodo.frx":1573
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrintBoletaxPeriodo.frx":1AB7
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrintBoletaxPeriodo.frx":1FFB
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrintBoletaxPeriodo.frx":210F
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrintBoletaxPeriodo.frx":2223
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrintBoletaxPeriodo.frx":2677
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrintBoletaxPeriodo.frx":27E3
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrintBoletaxPeriodo.frx":2D2B
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPrintBoletaxPeriodo.frx":3045
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmPrintBoletaxPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean
Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE
Dim Agregando As Boolean

Dim mMesActivo As Integer '--indica el mes activo

Private Sub CambiarMes()
    mMesActivo = SeleccionaMes(xCon)
    lblperiodo.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    Fg(0).Rows = Fg(0).FixedRows
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
    nTitulo = "Resumen de Planillas"
    nPeriodo = "Periodo: " + lblperiodo.Caption
    If NulosN(lbl_cod(0).Caption) <> 0 Then
         nTitulo1 = "Personal: " & StrConv(lbl_cb(0).Caption, 3)
    End If

    Me.MousePointer = vbHourglass
    oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg(mIndex), nTitulo, nPeriodo, nTitulo1, "Resumen de Planillas"
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Exportar"
End Sub


Private Sub pImprimirRes()
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
    nTitulo = "Resumen de Planillas"
    nPeriodo = "Periodo: " + lblperiodo.Caption
    nTitulo1 = "Proceso: " & StrConv(lbl_cb(0).Caption, 3) & ";    Categoría: " & StrConv(lbl_cb(2).Caption, 3)
    
    '--poniendo los resumenes
    Fg(0).Rows = Fg(0).Rows + 1
    Fg(0).TextMatrix(Fg(0).Rows - 1, 14) = "Totales >>"
    Fg(0).TextMatrix(Fg(0).Rows - 1, 15) = Format(GRID_SUMAR_COL(Fg(0), 15), FORMAT_MONTO)
    Fg(0).TextMatrix(Fg(0).Rows - 1, 16) = Format(GRID_SUMAR_COL(Fg(0), 16), FORMAT_MONTO)
    Fg(0).TextMatrix(Fg(0).Rows - 1, 17) = Format(GRID_SUMAR_COL(Fg(0), 17), FORMAT_MONTO)
    Fg(0).TextMatrix(Fg(0).Rows - 1, 18) = Format(GRID_SUMAR_COL(Fg(0), 18), FORMAT_MONTO)
    
    
    '--aplicando el fonfo
    GRID_COLOR_FONDO Fg(0), Fg(0).FixedRows, 15, Fg(0).Rows - 1, 15, &HE7FEFC
    GRID_COLOR_FONDO Fg(0), Fg(0).FixedRows, 16, Fg(0).Rows - 1, 16, &HC0E0FF
    GRID_COLOR_FONDO Fg(0), Fg(0).FixedRows, 17, Fg(0).Rows - 1, 17, &HC0C0FF
    GRID_COLOR_FONDO Fg(0), Fg(0).FixedRows, 18, Fg(0).Rows - 1, 18, &HFFD3A8
    
    
    Me.MousePointer = vbHourglass
    oPrint.Imprimir_x_VSFlexGrid Fg(mIndex), nTitulo, nTitulo1, nPeriodo, False, True
    Fg(0).Rows = Fg(0).Rows - 1
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


Private Sub Fg_DblClick(Index As Integer)
    If Fg(0).Row < 1 Then Exit Sub
    If Fg(0).Rows <= Fg(0).Row Then Exit Sub
    
    Fg(0).TextMatrix(Fg(0).Row, 4) = Not (Fg(0).TextMatrix(Fg(0).Row, 4))
End Sub

Private Sub fg_EnterCell(Index As Integer)
    If Fg(Index).Col <> 4 Then
        Fg(Index).Editable = flexEDNone
    Else
        Fg(Index).Editable = flexEDKbdMouse
    End If
End Sub

Private Sub Fg_KeyPress(Index As Integer, KeyAscii As Integer)
    If Fg(0).Row < 1 Then Exit Sub
    If Fg(0).Rows <= Fg(0).Row Then Exit Sub
    If KeyAscii = 32 Then Fg(0).TextMatrix(Fg(0).Row, 4) = Not (Fg(0).TextMatrix(Fg(0).Row, 4))


End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        mMesActivo = xMes
        
        SeEjecuto = True
        lblperiodo.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
        DoEvents
        pConfigurarGrilla
        
        LimpiaText txt_cb
        
        habilitar_Locked txt_cb, False
        DoEvents
        '--valores por defecto
        
        txt_cb(0).Text = 4: txt_cb_Validate 0, False
        txt_cb(2).Text = 1: txt_cb_Validate 2, False
       
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

Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 3 Then CambiarMes
    If Button.Index = 5 Then pExportar
    If Button.Index = 6 Then pImprimirDet
    If Button.Index = 8 Then
        Unload Me
    End If
End Sub

'****************************************************************************************

Private Sub cb_Click(Index As Integer)
    Dim xCampos() As String
    Dim nCampoBusca As String
    Dim nSQL As String
    Dim nTitulo As String
    On Error GoTo error

    
    Select Case Index
        Case 0 '--Proceso
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"

            nTitulo = "Seleccionar el Proceso"
            nSQL = "SELECT pla_proceso.id, pla_proceso.descripcion AS nombre, pla_proceso.id AS cod " _
                + vbCr + " FROM pla_proceso " _
                + vbCr + " WHERE (((pla_proceso.enproceso)=-1)); "

        Case 1 '--numero proceso menos el mensual
            
        Case 2 '--auto categoria
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Id":       xCampos(1, 1) = "id":        xCampos(1, 2) = "500":    xCampos(1, 3) = "N"
            nTitulo = "Buscando Categoría"
            nSQL = "SELECT mae_categoria.id, mae_categoria.descripcion AS nombre, mae_categoria.id AS cod " _
                + vbCr + " FROM mae_categoria;"

        Case 3 '--personal
            If NulosN(lbl_cod(2).Caption) = 0 Then
                MsgBox "Seleccione una Categoría", vbExclamation, xTitulo
                txt_cb(2).SetFocus
                Exit Sub
            End If
        
            ReDim xCampos(3, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "6500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Abrev":   xCampos(1, 1) = "abrev":    xCampos(1, 2) = "700":   xCampos(1, 3) = "C"
            xCampos(2, 0) = "Id":       xCampos(2, 1) = "id":        xCampos(2, 2) = "500":    xCampos(2, 3) = "N"

            nTitulo = "Seleccionar el Documento"
            nSQL = "SELECT mae_documento.id, mae_documento.descripcion AS nombre, mae_documento.id AS cod, mae_documento.abrev, mae_documento.codsun, mae_documentocta.idcuen " _
                + vbCr + " FROM mae_documento INNER JOIN mae_documentocta ON mae_documento.id = mae_documentocta.iddoc " _
                + vbCr + " WHERE (((mae_documentocta.idmon)=" & NulosN(lbl_cod(2).Caption) & ") AND ((mae_documentocta.tipope)=0));"

    End Select


    Dim xRs As New ADODB.Recordset
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio

    If xRs.State = 0 Then GoTo salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo salir

    lbl_cod(Index).Tag = lbl_cod(Index).Caption

    txt_cb(Index) = xRs.Fields(0) & "" '--TEXTO A MOSTRAR
    lbl_cb(Index).Caption = xRs.Fields(1) & "" '--NOMBRE
    lbl_cod(Index).Caption = xRs.Fields(2) & "" '--CODIGO
    lbl_cb(Index).ToolTipText = xRs.Fields(1) & "" '--NOMBRE
    
        
    If Trim(lbl_cod(Index).Tag) <> Trim(lbl_cod(Index).Caption) Then
        Select Case Index
            Case 0 '--auto tipo doc
                Fg(0).Rows = Fg(0).FixedRows
        End Select
    End If
    Select Case Index
        Case 0 '--proceso
'            txt_cb(4).SetFocus '--numero proceso (para obtener las horas)
        Case 1 '--auto categoria
            txt_cb(0).SetFocus '--proceso
        Case 2 '--moneda
        Case 4 '--numero proceso
            txt_cb(2).SetFocus '--moneda
    End Select

salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then pImprimirDet
    If ButtonMenu.Index = 2 Then pImprimirRes

End Sub

Private Sub txt_cb_Change(Index As Integer)
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cod(Index).Caption = ""
        Me.lbl_cb(Index).Tag = ""
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
        Case 0 '--proceso =>> ::auto
            nSQL = "SELECT pla_proceso.id, pla_proceso.descripcion AS nombre, pla_proceso.id AS cod " _
                + vbCr + " FROM pla_proceso " _
                + vbCr + " WHERE (((pla_proceso.enproceso)=-1)) and pla_proceso.id  = " & NulosN(txt_cb(Index).Text) & ""
        
        Case 1 '--numero
                
        Case 2 '--categoria
            nSQL = "SELECT mae_categoria.id, mae_categoria.descripcion AS nombre, mae_categoria.id AS cod " _
                + vbCr + " FROM mae_categoria " _
                + vbCr + " WHERE mae_categoria.id  = " & NulosN(txt_cb(Index).Text) & ""
            
        Case 3 '--personal
        
            Exit Sub
            
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
    
    If NulosN(lbl_cod(0).Caption) = 0 Then
        MsgBox "Seleccione el tipo de Proceso", vbExclamation, xTitulo
        txt_cb(0).SetFocus
        Exit Sub
    End If
    
    If NulosN(lbl_cod(2).Caption) = 0 Then
        MsgBox "Seleccione el tipo de Categoría", vbExclamation, xTitulo
        txt_cb(2).SetFocus
        Exit Sub
    End If
    
    If NulosN(lbl_cod(1).Caption) <> 0 Then
        nSQLIdEmp = " and pla_boleta.idemp = " & NulosN(lbl_cod(0).Caption)
        Exit Sub
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
'--esta consulta es la union de la consulta de empleado + la categoria + la boleta
    nSQL = "SELECT  *, " _
        + vbCr + " (SELECT Sum(pla_marcacionhora.totseg) AS totseg FROM mae_tipohora INNER JOIN (pla_marcacion INNER JOIN pla_marcacionhora ON pla_marcacion.id = pla_marcacionhora.idmarca) ON mae_tipohora.id = pla_marcacionhora.idhora WHERE (((pla_marcacionhora.idemp)=emp.idemp) AND ((Year([pla_marcacion].[dia]))=" & AnoTra & ") AND ((Month([pla_marcacion].[dia]))=" & mMesActivo & ") AND ((mae_tipohora.hortrabajo)=-1)) GROUP BY pla_marcacionhora.idemp) AS totseg " _
        + vbCr + " FROM (SELECT * FROM " _
        + vbCr + " (SELECT pla_empleados.id AS idemp, mae_dociden.abrev AS docabrev, pla_empleados.numdoc AS docemp, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, pla_empleados.fchnac, mae_sexo.abrev AS sexo, pla_empleados.idcargo, mae_cargo.descripcion AS cargo " _
        + vbCr + " FROM mae_sexo RIGHT JOIN (mae_cargo RIGHT JOIN (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) ON mae_cargo.id = pla_empleados.idcargo) ON mae_sexo.id = pla_empleados.idsex " _
        + vbCr + " WHERE (((pla_empleados.idbolpag)=" & NulosN(lbl_cod(0).Caption) & "))" _
        + vbCr + " ORDER BY [pla_empleados].[nom] & ' ' & [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat]) AS emp " _
        + vbCr + " INNER JOIN " _
        + vbCr + " (SELECT pla_periodolaboral.idemp AS idemp1, mae_categoria.descripcion AS categoria, mae_categoria.nomcor AS catabrev, Last(pla_periodolaboral.fchini) AS ingreso " _
        + vbCr + " FROM mae_categoria INNER JOIN pla_periodolaboral ON mae_categoria.id = pla_periodolaboral.idcat " _
        + vbCr + " Where   pla_periodolaboral.idcat=" & NulosN(lbl_cod(2).Caption) & " " _
        + vbCr + " GROUP BY pla_periodolaboral.idemp, mae_categoria.descripcion, mae_categoria.nomcor " _
        + vbCr + " ORDER BY pla_periodolaboral.idemp, Last(pla_periodolaboral.fchini), Last(pla_periodolaboral.fchfin)) AS periodo " _
        + vbCr + " ON emp.idemp = periodo.idemp1) AS emp " _
        + vbCr + " INNER JOIN " _
        + vbCr + " (SELECT pla_boleta.id AS idbol,  pla_boleta.idemp as idemp1,   pla_boleta.numreg, pla_boleta.idmon, pla_boleta.numser, pla_boleta.numdoc, pla_boleta.fchdoc, pla_boleta.fchpago, mae_moneda.simbolo, pla_boleta.impingr, pla_boleta.impapor, pla_boleta.impdesc, pla_boleta.imptot " _
        + vbCr + " FROM pla_proceso RIGHT JOIN (mae_moneda RIGHT JOIN pla_boleta ON mae_moneda.id = pla_boleta.idmon) ON pla_proceso.id = pla_boleta.idproc " _
        + vbCr + " WHERE pla_boleta.ano= " & AnoTra & " and pla_boleta.idmes= " & mMesActivo & " and pla_boleta.idproc= " & NulosN(lbl_cod(0).Caption) & ") AS boleta ON emp.idemp = boleta.idemp1 " _
        + vbCr + " ORDER BY emp.nombres"
        '(((pla_periodolaboral.fchfin) Is Null)) AND
    RST_Busq RstTmp, nSQL, xCon
    
    Fg(0).Rows = Fg(0).FixedRows
    If RstTmp.State = 0 Then GoTo salir

    Agregando = True


    '--de las horas del mes
    '--OPCIONAL CAMBIAR LUEGO
    Dim mTotalSegundosMes As Long
    Dim mTotalDias As Integer
    mTotalDias = HallaDiasMes(CDate("01/" & mMesActivo & "/" & AnoTra))
    mTotalSegundosMes = mTotalDias * 8
    mTotalSegundosMes = mTotalSegundosMes * 60 * 60
    '--FIN OPCIONAL CAMBIAR LUEGO

    If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    Do While Not RstTmp.EOF
        Fg(0).Rows = Fg(0).Rows + 1
        Fg(0).TextMatrix(Fg(0).Rows - 1, 1) = NulosN(RstTmp("idbol"))
        
        Fg(0).TextMatrix(Fg(0).Rows - 1, 4) = -1 '--seleccion por defecto
        
        Fg(0).TextMatrix(Fg(0).Rows - 1, 2) = NulosC(RstTmp("idemp"))
        Fg(0).TextMatrix(Fg(0).Rows - 1, 3) = NulosC(RstTmp("idmon"))
        
        Fg(0).TextMatrix(Fg(0).Rows - 1, 5) = NulosC(RstTmp("nombres"))
        Fg(0).TextMatrix(Fg(0).Rows - 1, 6) = NulosC(RstTmp("catabrev"))
        Fg(0).TextMatrix(Fg(0).Rows - 1, 7) = NulosC(RstTmp("cargo"))
        Fg(0).TextMatrix(Fg(0).Rows - 1, 8) = Format(NulosC(RstTmp("ingreso")), FORMAT_DATE)
        
        Fg(0).TextMatrix(Fg(0).Rows - 1, 9) = ConvertHora(mTotalSegundosMes) ' REEMPLAZAR LUEGO NulosN(RstTmp("totseg"))
        
        Fg(0).TextMatrix(Fg(0).Rows - 1, 10) = Format(NulosC(RstTmp("fchdoc")), FORMAT_DATE)
        Fg(0).TextMatrix(Fg(0).Rows - 1, 11) = Format(NulosC(RstTmp("fchpago")), FORMAT_DATE)
        
        Fg(0).TextMatrix(Fg(0).Rows - 1, 12) = NulosC(RstTmp("simbolo"))
        
        Fg(0).TextMatrix(Fg(0).Rows - 1, 13) = NulosC(RstTmp("numser"))
        Fg(0).TextMatrix(Fg(0).Rows - 1, 14) = NulosC(RstTmp("numdoc"))
                
        Fg(0).TextMatrix(Fg(0).Rows - 1, 15) = Format(RstTmp("impingr"), FORMAT_MONTO)
        Fg(0).TextMatrix(Fg(0).Rows - 1, 16) = Format(RstTmp("impdesc"), FORMAT_MONTO)
        Fg(0).TextMatrix(Fg(0).Rows - 1, 17) = Format(RstTmp("impapor"), FORMAT_MONTO)
        Fg(0).TextMatrix(Fg(0).Rows - 1, 18) = Format(RstTmp("imptot"), FORMAT_MONTO)
        
        RstTmp.MoveNext
    Loop
    '--aplicando el fonfo
    GRID_COLOR_FONDO Fg(0), Fg(0).FixedRows, 15, Fg(0).Rows - 1, 15, &HE7FEFC
    GRID_COLOR_FONDO Fg(0), Fg(0).FixedRows, 16, Fg(0).Rows - 1, 16, &HC0E0FF
    GRID_COLOR_FONDO Fg(0), Fg(0).FixedRows, 17, Fg(0).Rows - 1, 17, &HC0C0FF
    GRID_COLOR_FONDO Fg(0), Fg(0).FixedRows, 18, Fg(0).Rows - 1, 18, &HFFD3A8

salir:
    Set RstTmp = Nothing
    Agregando = False


End Sub

Private Sub pConfigurarGrilla()
    Agregando = True

    With Fg(0) '--proceso autimatico
        .Clear
        .Rows = 2
        .FixedRows = 2
        .Cols = 19
        .RowHeight(0) = 250
        .RowHeight(1) = 500
        .FrozenCols = 5

        GRID_COMBINAR Fg(0), 0, 1, 0, 3, "id's", flexAlignCenterCenter, True, flexMergeFree, vbBlack, &HD8E9EC, True
        GRID_COMBINAR Fg(0), 0, 4, 0, 11, " ", flexAlignCenterCenter, True, flexMergeFree, vbBlack, &HD8E9EC, True
        GRID_COMBINAR Fg(0), 0, 13, 0, 14, "Documento", flexAlignCenterCenter, True, flexMergeFree, vbBlack, &HD8E9EC, True
        GRID_COMBINAR Fg(0), 0, 15, 0, 17, "Totales", flexAlignCenterCenter, True, flexMergeFree, vbBlack, &HD8E9EC, True
        GRID_COMBINAR Fg(0), 0, 18, 0, 18, " ", flexAlignCenterCenter, True, flexMergeFree, vbBlack, &HD8E9EC, True

        .TextMatrix(1, 1) = "IdBol":        .ColWidth(1) = 0:
        .TextMatrix(1, 2) = "IdEmp":        .ColWidth(2) = 0:
        .TextMatrix(1, 3) = "IdMon":        .ColWidth(3) = 0:

        .TextMatrix(1, 4) = "Sel":          .ColWidth(4) = 350:    .ColAlignment(4) = flexAlignCenterCenter: .Row = 1: .Col = 4: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 5) = "Personal":     .ColWidth(5) = 2400:   .ColAlignment(5) = flexAlignLeftCenter:   .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 6) = "Cat.":         .ColWidth(6) = 450:    .ColAlignment(6) = flexAlignLeftCenter:   .Row = 1: .Col = 6: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 7) = "Cargo":        .ColWidth(7) = 790:    .ColAlignment(7) = flexAlignLeftCenter:   .Row = 1: .Col = 7: .CellAlignment = flexAlignLeftCenter

        .TextMatrix(1, 8) = "Fecha" & vbCr & "Ingreso":   .ColWidth(8) = 800:    .ColAlignment(8) = flexAlignCenterCenter: .Row = 1: .Col = 8: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 9) = "Horas" & vbCr & "Trabajo":   .ColWidth(9) = 790:    .ColAlignment(9) = flexAlignRightCenter:  .Row = 1: .Col = 9: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 10) = "Fecha" & vbCr & "Emisión":  .ColWidth(10) = 800:   .ColAlignment(10) = flexAlignCenterCenter: .Row = 1: .Col = 10: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 11) = "Fecha" & vbCr & "Pago":     .ColWidth(11) = 800:   .ColAlignment(11) = flexAlignCenterCenter: .Row = 1: .Col = 11: .CellAlignment = flexAlignCenterCenter

        .TextMatrix(1, 12) = "M":          .ColWidth(12) = 450:  .ColAlignment(12) = flexAlignCenterCenter: .Row = 1: .Col = 12: .CellAlignment = flexAlignCenterCenter
        
        .TextMatrix(1, 13) = "Serie":      .ColWidth(13) = 450:  .ColAlignment(13) = flexAlignCenterCenter: .Row = 1: .Col = 13: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 14) = "Número":     .ColWidth(14) = 1000:  .ColAlignment(14) = flexAlignCenterCenter: .Row = 1: .Col = 14: .CellAlignment = flexAlignCenterCenter

        .TextMatrix(1, 15) = "Ingresos":    .ColWidth(15) = 700:  .ColAlignment(15) = flexAlignRightCenter:  .Row = 1: .Col = 15: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 16) = "Descuento":   .ColWidth(16) = 850:  .ColAlignment(16) = flexAlignRightCenter:  .Row = 1: .Col = 16: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 17) = "Aportes":     .ColWidth(17) = 700:  .ColAlignment(17) = flexAlignRightCenter:  .Row = 1: .Col = 17: .CellAlignment = flexAlignRightCenter

        .TextMatrix(1, 18) = "Neto a" & vbCr & "Pagar":  .ColWidth(18) = 850:  .ColAlignment(18) = flexAlignRightCenter:  .Row = 1: .Col = 18: .CellAlignment = flexAlignRightCenter


        .ColDataType(4) = flexDTBoolean
        .SelectionMode = flexSelectionByRow

    End With
    Agregando = False
    DoEvents
End Sub



'*************Imprimir

Private Sub pImprimirDet()
    '===================================================================================================
    'Creado : //08 Por: Johan Castro
    'Propósito: Imprimir la boleta de pago
    '
    'Entradas:  Ninguna
    '
    'Resultados: Reporte de la boleta de pago lista para impresion
    '
    'Modificado: 05/01/11 Por: Johan Castro
    '           1.- Quitar los valores por defecto que hagan referencia al periodo laboral
    '           para que se muestre tienen que ingresarse manualmente en resumen de horas
    '           antes de registrar la boleta
    '           2.- Agregar campo de tipo de documento para impresion de boleta
    '===================================================================================================

    On Error GoTo error
    Dim RstEmp As New ADODB.Recordset
    Dim mTotalDias As Integer
    Dim nTotalHN As String
    Dim nTotalHE1 As String
    Dim nTotalHE2 As String
    Dim nSQL As String
    Dim mIdEmp&
    
    Dim mCuenta&
    Dim K&
    
        
    mCuenta = -1
    For K = Fg(0).FixedRows To Fg(0).Rows - 1
        If NulosN(Fg(0).TextMatrix(K, 4)) = -1 Then mCuenta = mCuenta + 1
    Next K
    If mCuenta = -1 Then
        MsgBox "Seleccione los registos que desea imprimir", vbExclamation, xTitulo
        Exit Sub
    End If
    
    '--de las horas
'    mTotalDias = HallaDiasMes(CDate("01/" & mMesActivo & "/" & AnoTra))
    
    '--de las horas del mes
    Dim mTotalSegundosMes As Long
'    mTotalSegundosMes = mTotalDias * 8
'    mTotalSegundosMes = mTotalSegundosMes * 60 * 60
'    nTotalHN = "'" & ConvertHora(mTotalSegundosMes) & "'"
'
'    nTotalHE1 = "'00:00:00'"
'    nTotalHE2 = "'00:00:00'"
    
    mTotalDias = 0
    mTotalSegundosMes = 0
    mTotalSegundosMes = 0
    nTotalHN = "'00:00:00'"
    nTotalHE1 = "'00:00:00'"
    nTotalHE2 = "'00:00:00'"
    
    Dim nSQLIdBol As String
    Dim RstIngreso As New ADODB.Recordset
    Dim RstDescuento As New ADODB.Recordset
    Dim RstAportacion As New ADODB.Recordset
    
    Me.MousePointer = vbHourglass
    
    nSQLIdBol = GRID_GENERAR_SQL_ID(Fg(0), 1, "pla_boleta.id", "IN", True, 4, -1)

    DoEvents
    
    nSQL = "SELECT UCase(Left([con_meses].[descripcion],3)) & ' ' & [pla_boleta].[ano] AS periodo, mae_area.descripcion AS area, mae_cargo.descripcion AS cargo, pla_empleados.id AS idemp, pla_empleados!apepat+' '+pla_empleados!apemat+', '+pla_empleados!nom AS apenom, pla_empleados.numdoc, pla_categoria1.cuspp, pla_empleados.numessalud, pla_empleados.basico, [pla_boleta].[numser] & ' ' & [pla_boleta].[numdoc] AS numboleta,  iif(pla_marcaresumemp.totdiatra is null, " & mTotalDias & ",pla_marcaresumemp.totdiatra)  AS DiaTrabajo, iif( pla_marcaresumhor.tothor is null," & nTotalHN & ", pla_marcaresumhor.tothor) AS TotalHN," & nTotalHE1 & " AS TotalHE1, " & nTotalHE2 & " AS TotalHE2, pla_empleados.fching AS fchingreso, pla_empleados.fchcese, pla_boleta.idmon,ucase(mae_documento.descripcion) AS tipodoc " _
        + vbCr + " FROM (((mae_area RIGHT JOIN (pla_empleados LEFT JOIN mae_cargo ON pla_empleados.idcargo = mae_cargo.id) ON mae_area.id = pla_empleados.idarea) LEFT JOIN pla_categoria1 ON pla_empleados.id = pla_categoria1.idemp) INNER JOIN (((pla_boleta INNER JOIN con_meses ON pla_boleta.idmes = con_meses.id) LEFT JOIN pla_marcaresumemp ON (pla_boleta.idresmarca = pla_marcaresumemp.idres) AND (pla_boleta.idemp = pla_marcaresumemp.idemp)) LEFT JOIN pla_marcaresumhor ON (pla_boleta.idresmarca = pla_marcaresumhor.idres) AND (pla_boleta.idemp = pla_marcaresumhor.idemp)) ON pla_empleados.id = pla_boleta.idemp) LEFT JOIN mae_documento ON pla_boleta.iddoc = mae_documento.id " _
        + vbCr + " WHERE (pla_marcaresumhor.idhora = 15 or pla_marcaresumhor.idhora is null ) and " & nSQLIdBol & " " _
        + vbCr + " GROUP BY UCase(Left([con_meses].[descripcion],3)) & ' ' & [pla_boleta].[ano], mae_area.descripcion, mae_cargo.descripcion, pla_empleados.id, pla_empleados!apepat+' '+pla_empleados!apemat+', '+pla_empleados!nom, pla_empleados.numdoc, pla_categoria1.cuspp, pla_empleados.numessalud, pla_empleados.basico, [pla_boleta].[numser] & ' ' & [pla_boleta].[numdoc], pla_boleta.idmon,iif(pla_marcaresumemp.totdiatra is null, " & mTotalDias & ",pla_marcaresumemp.totdiatra),iif( pla_marcaresumhor.tothor is null," & nTotalHN & ", pla_marcaresumhor.tothor),pla_empleados.fching , pla_empleados.fchcese,pla_empleados.apepat,mae_documento.descripcion " _
        + vbCr + " ORDER BY pla_empleados.apepat"


'    nSQL = "SELECT " & nPeriodo & " AS periodo, mae_area.descripcion AS area, mae_cargo.descripcion AS cargo,pla_empleados.id as idemp, pla_empleados!apepat+' '+pla_empleados!apemat+', '+pla_empleados!nom AS apenom, pla_empleados.numdoc, pla_categoria1.cuspp, pla_empleados.numessalud, pla_empleados.basico, " & nNumBoleta & " AS numboleta, " & mTotalDias & " AS DiaTrabajo, " & nTotalHN & " AS TotalHN," & nTotalHE1 & " AS TotalHE1, " & nTotalHE2 & " AS TotalHE2, Last(pla_periodolaboral.fchini) AS fchingreso, Last(pla_periodolaboral.fchfin) AS fchcese, " & mIdMon & " AS idmon " _
'        + vbCr + " FROM ((mae_area RIGHT JOIN (pla_empleados LEFT JOIN mae_cargo ON pla_empleados.idcargo = mae_cargo.id) ON mae_area.id = pla_empleados.idarea) LEFT JOIN pla_categoria1 ON pla_empleados.id = pla_categoria1.idemp) LEFT JOIN pla_periodolaboral ON pla_empleados.id = pla_periodolaboral.idemp " _
'        + vbCr + " Group By  mae_area.descripcion, mae_cargo.descripcion, pla_empleados!apepat+' '+pla_empleados!apemat+', '+pla_empleados!nom, pla_empleados.numdoc, pla_categoria1.cuspp, pla_empleados.numessalud, pla_empleados.basico, pla_empleados.id " _
'        + vbCr + " Having (((pla_empleados.id) = " & mIdEmp & " )) " _
'        + vbCr + " ORDER BY Last(pla_periodolaboral.fchini), Last(pla_periodolaboral.fchfin); "


    RST_Busq RstEmp, nSQL, xCon
    DoEvents
    pConceptoDocumentoEmp RstIngreso, nSQLIdBol, e_Remuneracion, True
    DoEvents
    pConceptoDocumentoEmp RstDescuento, nSQLIdBol, e_Descuento, True
    DoEvents
    pConceptoDocumentoEmp RstAportacion, nSQLIdBol, e_Aportacion, True
    DoEvents
    FrmPrintBoleta.pRecibeRsts RstIngreso, RstDescuento, RstAportacion, RstEmp
    FrmPrintBoleta.Show
    
    Set RstEmp = Nothing
    Set RstIngreso = Nothing
    Set RstDescuento = Nothing
    Set RstAportacion = Nothing
    Me.MousePointer = vbDefault
    
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimirDet"
    Set RstEmp = Nothing
    Set RstIngreso = Nothing
    Set RstDescuento = Nothing
    Set RstAportacion = Nothing
    Me.MousePointer = vbDefault
    
End Sub


