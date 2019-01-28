VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmActDatosContables 
   Caption         =   "Form1"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7200
      Left            =   15
      TabIndex        =   0
      Top             =   375
      Width           =   12000
      _cx             =   21167
      _cy             =   12700
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
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
      Caption         =   "   Consulta   "
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6780
         Left            =   45
         TabIndex        =   1
         Top             =   375
         Width           =   11910
         Begin VB.Frame Frame2 
            Caption         =   "( Detalle del Item )"
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
            Height          =   1335
            Left            =   30
            TabIndex        =   5
            Top             =   5445
            Width           =   11865
            Begin VB.TextBox TxtDescCosto 
               Height          =   300
               Left            =   1380
               TabIndex        =   15
               Text            =   "TxtDescCosto"
               Top             =   615
               Width           =   4395
            End
            Begin VB.TextBox TxtDescCuenta 
               Height          =   300
               Left            =   1380
               TabIndex        =   12
               Text            =   "TxtDescCuenta"
               Top             =   300
               Width           =   4395
            End
            Begin VB.TextBox TxtRetencion 
               Height          =   300
               Left            =   7395
               TabIndex        =   10
               Text            =   "TxtRetencion"
               Top             =   930
               Width           =   4395
            End
            Begin VB.TextBox TxtPercepcion 
               Height          =   300
               Left            =   7395
               TabIndex        =   8
               Text            =   "TxtPercepcion"
               Top             =   615
               Width           =   4395
            End
            Begin VB.TextBox TxtDetraccion 
               Height          =   300
               Left            =   7395
               TabIndex        =   6
               Text            =   "TxtDetraccion"
               Top             =   300
               Width           =   4395
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Centro de Costo"
               Height          =   195
               Index           =   4
               Left            =   135
               TabIndex        =   14
               Top             =   645
               Width           =   1140
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Cuenta"
               Height          =   195
               Index           =   3
               Left            =   135
               TabIndex        =   13
               Top             =   345
               Width           =   510
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Percepcion"
               Height          =   195
               Index           =   2
               Left            =   6255
               TabIndex        =   11
               Top             =   945
               Width           =   810
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Retencion "
               Height          =   195
               Index           =   1
               Left            =   6255
               TabIndex        =   9
               Top             =   645
               Width           =   780
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Detraccion"
               Height          =   195
               Index           =   0
               Left            =   6255
               TabIndex        =   7
               Top             =   345
               Width           =   780
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   5100
            Left            =   30
            TabIndex        =   2
            Top             =   345
            Width           =   11865
            _cx             =   20929
            _cy             =   8996
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
            Cols            =   15
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmActDatosContables.frx":0000
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
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Items de Almacen"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   90
            TabIndex        =   4
            Top             =   45
            Width           =   11715
         End
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
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActDatosContables.frx":01BD
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActDatosContables.frx":0701
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActDatosContables.frx":0A93
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActDatosContables.frx":0C17
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActDatosContables.frx":106B
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActDatosContables.frx":1183
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActDatosContables.frx":16C7
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActDatosContables.frx":1C0B
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActDatosContables.frx":1D1F
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActDatosContables.frx":1E33
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActDatosContables.frx":2287
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmActDatosContables.frx":23F3
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Guia"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmActDatosContables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

Dim RstItem As New ADODB.Recordset
Dim SeEjecuto As Boolean

Private Sub Fg1_RowColChange()
    If Fg1.Rows = 0 Then Exit Sub
    MuestraDatos
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        Dim A As Integer
        RST_Busq RstItem, "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, " _
            & " alm_inventario.idret, alm_inventario.idper, alm_inventario.iddet, alm_inventario.idcuenta, alm_inventario.idcencos, " _
            & " con_planctas.cuenta, con_planctas.descripcion AS desccuen, mae_percepcion.descripcion AS descper, " _
            & " mae_retencion.descripcion AS descret, mae_detraccion.descripcion AS descdet, con_centrocosto.codigo AS codcencos, " _
            & " con_centrocosto.descripcion AS desccencos, mae_tipoproducto.descripcion AS desctippro " _
            & " FROM mae_tipoproducto RIGHT JOIN (mae_unidades RIGHT JOIN (((((mae_percepcion RIGHT JOIN alm_inventario " _
            & " ON mae_percepcion.id = alm_inventario.idper) LEFT JOIN con_planctas ON alm_inventario.idcuenta = con_planctas.id) " _
            & " LEFT JOIN mae_detraccion ON alm_inventario.iddet = mae_detraccion.id) LEFT JOIN mae_retencion ON alm_inventario.idret = mae_retencion.id) " _
            & " LEFT JOIN con_centrocosto ON alm_inventario.idcencos = con_centrocosto.id) ON mae_unidades.id = alm_inventario.idunimed) " _
            & " ON mae_tipoproducto.id = alm_inventario.tippro", xCon

        If RstItem.RecordCount <> 0 Then
            Fg1.Rows = 1
            RstItem.MoveFirst
            For A = 1 To RstItem.RecordCount
                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(A, 1) = RstItem("codpro")
                Fg1.TextMatrix(A, 2) = RstItem("descripcion")
                Fg1.TextMatrix(A, 3) = RstItem("abrev")
                Fg1.TextMatrix(A, 4) = RstItem("cuenta")
                Fg1.TextMatrix(A, 5) = RstItem("iddet")
                Fg1.TextMatrix(A, 6) = RstItem("idret")
                Fg1.TextMatrix(A, 7) = RstItem("idper")
                Fg1.TextMatrix(A, 8) = NulosC(RstItem("codcencos"))
                Fg1.TextMatrix(A, 9) = RstItem("id")
                
                Fg1.TextMatrix(A, 10) = NulosC(RstItem("descdet"))
                Fg1.TextMatrix(A, 11) = NulosC(RstItem("descret"))
                Fg1.TextMatrix(A, 12) = NulosC(RstItem("descper"))
                Fg1.TextMatrix(A, 13) = NulosC(RstItem("desccencos"))
                Fg1.TextMatrix(A, 14) = NulosC(RstItem("desccuen"))
                
                RstItem.MoveNext
                If RstItem.EOF = True Then
                    Exit For
                End If
            Next A
        End If
    End If
End Sub

Sub MuestraDatos()
    TxtDescCuenta.Text = NulosC(Fg1.TextMatrix(Fg1.Row, 14))
    TxtDescCosto.Text = NulosC(Fg1.TextMatrix(Fg1.Row, 13))
    TxtDetraccion.Text = NulosC(Fg1.TextMatrix(Fg1.Row, 10))
    TxtPercepcion.Text = NulosC(Fg1.TextMatrix(Fg1.Row, 11))
    TxtRetencion.Text = NulosC(Fg1.TextMatrix(Fg1.Row, 12))
End Sub

Private Sub Form_Load()
    Fg1.ColWidth(9) = 0
    Fg1.ColWidth(10) = 0
    Fg1.ColWidth(11) = 0
    Fg1.ColWidth(12) = 0
    Fg1.ColWidth(13) = 0
    Fg1.ColWidth(14) = 0
End Sub

