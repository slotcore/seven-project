VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmManDiaFestivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Dias Festivos"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   11865
   Begin VB.Frame FraEditor 
      BorderStyle     =   0  'None
      Height          =   3270
      Left            =   3435
      TabIndex        =   15
      Top             =   1410
      Visible         =   0   'False
      Width           =   5430
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   5160
         Picture         =   "FrmManDiaFestivo.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   27
         ToolTipText     =   "Cerrar"
         Top             =   75
         Width           =   195
      End
      Begin VB.TextBox txt 
         Height          =   870
         Index           =   2
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Tag             =   "null"
         Text            =   "FrmManDiaFestivo.frx":02EC
         Top             =   1740
         Width           =   5190
      End
      Begin VB.CommandButton CmdEditor 
         Caption         =   "&Grabar"
         Height          =   420
         Index           =   0
         Left            =   1530
         TabIndex        =   6
         Top             =   2760
         Width           =   1020
      End
      Begin VB.CommandButton CmdEditor 
         Caption         =   "&Cancelar"
         Height          =   420
         Index           =   1
         Left            =   2850
         TabIndex        =   7
         Top             =   2760
         Width           =   1020
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   1
         Left            =   990
         TabIndex        =   0
         Text            =   "txt(1)"
         Top             =   405
         Width           =   4335
      End
      Begin VB.CommandButton cb 
         Height          =   240
         Index           =   1
         Left            =   4695
         Picture         =   "FrmManDiaFestivo.frx":02F5
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1140
         Width           =   255
      End
      Begin VB.CommandButton cb 
         Height          =   240
         Index           =   0
         Left            =   4695
         Picture         =   "FrmManDiaFestivo.frx":0427
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   795
         Width           =   255
      End
      Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
         Height          =   315
         Index           =   0
         Left            =   1005
         TabIndex        =   1
         Top             =   765
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Valor           =   "  /  /    "
      End
      Begin AspaTextBoxFecha.TextBoxFecha txtfecha 
         Height          =   315
         Index           =   1
         Left            =   1005
         TabIndex        =   3
         Top             =   1110
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Valor           =   "  /  /    "
      End
      Begin MSComCtl2.DTPicker dtpk 
         Height          =   300
         Index           =   0
         Left            =   3540
         TabIndex        =   2
         Top             =   765
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   529
         _Version        =   393216
         Format          =   57475074
         CurrentDate     =   39534
      End
      Begin MSComCtl2.DTPicker dtpk 
         Height          =   300
         Index           =   1
         Left            =   3540
         TabIndex        =   4
         Top             =   1110
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   529
         _Version        =   393216
         Format          =   57475074
         CurrentDate     =   39534
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observación"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   25
         Top             =   1500
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   23
         Top             =   555
         Width           =   840
      End
      Begin VB.Label lbltxtfch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.Fin"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   22
         Top             =   1230
         Width           =   345
      End
      Begin VB.Label lbldtpk 
         AutoSize        =   -1  'True
         Caption         =   "H.Fin"
         Height          =   195
         Index           =   1
         Left            =   2895
         TabIndex        =   21
         Top             =   1230
         Width           =   375
      End
      Begin VB.Label lbldtpk 
         AutoSize        =   -1  'True
         Caption         =   "H.Inicio"
         Height          =   195
         Index           =   0
         Left            =   2895
         TabIndex        =   20
         Top             =   885
         Width           =   540
      End
      Begin VB.Label lbltxtfch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.Inicio"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   17
         Top             =   885
         Width           =   510
      End
      Begin VB.Label LblTituloFrame 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Editor de Dia Festivo"
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
         Left            =   90
         TabIndex        =   16
         Top             =   90
         Width           =   1800
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5415
         X2              =   5415
         Y1              =   -75
         Y2              =   4815
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -330
         X2              =   5715
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   15
         X2              =   5685
         Y1              =   3240
         Y2              =   3255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   15
         Y1              =   15
         Y2              =   5070
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         Index           =   2
         X1              =   90
         X2              =   5280
         Y1              =   2670
         Y2              =   2670
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   300
         Index           =   1
         Left            =   30
         Top             =   45
         Width           =   5355
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   15
      TabIndex        =   9
      Top             =   360
      Width           =   11835
      _cx             =   20876
      _cy             =   12726
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Style           =   3
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
         Height          =   6840
         Left            =   45
         TabIndex        =   10
         Top             =   330
         Width           =   11745
         Begin VB.Frame fra 
            Height          =   660
            Index           =   0
            Left            =   15
            TabIndex        =   11
            Top             =   5580
            Width           =   11670
            Begin VB.CommandButton cmd 
               Caption         =   "Eliminar"
               Height          =   345
               Index           =   2
               Left            =   2865
               TabIndex        =   14
               Top             =   210
               Width           =   1200
            End
            Begin VB.CommandButton cmd 
               Caption         =   "Agregar"
               Height          =   345
               Index           =   0
               Left            =   60
               TabIndex        =   13
               Top             =   210
               Width           =   1200
            End
            Begin VB.CommandButton cmd 
               Caption         =   "Modificar"
               Height          =   345
               Index           =   1
               Left            =   1305
               TabIndex        =   12
               Top             =   210
               Width           =   1200
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   5175
            Left            =   15
            TabIndex        =   24
            Top             =   345
            Width           =   11670
            _cx             =   20585
            _cy             =   9128
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
            FormatString    =   $"FrmManDiaFestivo.frx":0559
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
         Begin VB.Label lblperiodo 
            Alignment       =   1  'Right Justify
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
            Left            =   9690
            TabIndex        =   26
            Top             =   30
            Width           =   1980
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7065
      Top             =   -15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDiaFestivo.frx":0595
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDiaFestivo.frx":0AD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDiaFestivo.frx":0E6B
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDiaFestivo.frx":0FEF
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDiaFestivo.frx":1443
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDiaFestivo.frx":155B
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDiaFestivo.frx":1A9F
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDiaFestivo.frx":1FE3
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDiaFestivo.frx":20F7
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDiaFestivo.frx":220B
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDiaFestivo.frx":265F
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDiaFestivo.frx":27CB
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManDiaFestivo.frx":2D13
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Periodo"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Actualizar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "&Agregar"
      End
      Begin VB.Menu menu1_2 
         Caption         =   "&Modificar"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_4 
         Caption         =   "&Eliminar"
      End
   End
End
Attribute VB_Name = "FrmManDiaFestivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim Agregando As Boolean

Dim mMesActivo As Integer '--indica el mes activo

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0 '--Add dia festivo
            nuevo
        Case 1 '--Modificar dia festivo
            Modificar
        Case 2 '--Eliminar dia festivo
            Eliminar
    End Select
End Sub

Private Sub CmdEditor_Click(Index As Integer)
    Select Case Index
        Case 0 'grabar
            If Grabar() = True Then
                pCargarGrid
                If QueHace = 1 Then
                    If MsgBox("Desea Agregar otro Registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbYes Then
                        nuevo
                    Else
                        CmdEditor_Click 1
                    End If
                Else
                    CmdEditor_Click 1
                End If
            End If
        Case 1 'cancelar
            Cancelar
    End Select
End Sub

Private Sub dtpk_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If dtpk(Index).Enabled = False Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click Index
    ElseIf KeyCode = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub pCargarGrid()
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL  As String
    On Error GoTo error
    lblperiodo(0).Caption = "Periodo: " & AnoTra
    nSQL = "SELECT mae_diasfestivos.* " _
        + vbCr + " FROM mae_diasfestivos WHERE anno = " & AnoTra & " " _
        + vbCr + " ORDER BY mae_diasfestivos.fchini, mae_diasfestivos.horini; "

    Me.MousePointer = vbHourglass
    RST_Busq RstTmp, nSQL, xCon
    '---------------
    pConfigurarGrilla
    '---------------
    If RstTmp.RecordCount <> 0 Then
        Agregando = True
        If RstTmp.RecordCount <> 0 Then
            RstTmp.MoveFirst
            Do While Not RstTmp.EOF
                With Fg1
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = NulosN(RstTmp.Fields("id"))
                    .TextMatrix(.Rows - 1, 2) = NulosC(RstTmp.Fields("descripcion"))
                    .TextMatrix(.Rows - 1, 3) = NulosC(RstTmp.Fields("fchini"))
                    .TextMatrix(.Rows - 1, 4) = NulosC(RstTmp.Fields("horini"))
                    .TextMatrix(.Rows - 1, 5) = NulosC(RstTmp.Fields("fchfin"))
                    .TextMatrix(.Rows - 1, 6) = NulosC(RstTmp.Fields("horfin"))
                    .TextMatrix(.Rows - 1, 7) = NulosC(RstTmp.Fields("observacion"))
                    RstTmp.MoveNext
                End With
            Loop
        End If
    End If
    If Fg1.Rows > 1 Then
        Fg1.Row = Fg1.Rows - 1
        If Fg1.Enabled = True Then Fg1.SetFocus
    End If
    '---------------
    Me.MousePointer = vbDefault
    Exit Sub
error:
    SHOW_ERROR Me.Name, "pCargarGrid"
End Sub

Private Sub Fg1_DblClick()
    cmd_Click 1
End Sub

Private Sub Fg1_EnterCell()
    Fg1.Editable = flexEDNone
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If Fg1.Enabled = False Then Exit Sub
    If KeyCode = 45 Then
        nuevo
    End If
    If KeyCode = 46 Then
        Eliminar
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Fg1.Enabled = False Then Exit Sub
    If Button = 2 Then
        PopupMenu menu1
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    mMesActivo = xMes
    
    SeEjecuto = False
    pCargarGrid
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If FraEditor.Visible = True Then CmdEditor_Click 1
    End If
End Sub

Private Sub Form_Load()
    CentrarFrm Me
    SeEjecuto = False
    Agregando = False
    QueHace = 3
    
    TabOne1.CurrTab = 0
    Frame1.BackColor = &H8000000F
    '--
    txtfecha(0).Valor = Date
    txtfecha(1).Valor = Date
    '--
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation, xTitulo
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub Menu1_1_Click()
    cmd_Click 0
End Sub

Private Sub menu1_2_Click()
cmd_Click 1
End Sub

Private Sub Menu1_3_Click()
    cmd_Click 4
End Sub

Private Sub menu1_4_Click()
    cmd_Click 2
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then CambiarMes
    If Button.Index = 3 Then Buscar
    If Button.Index = 4 Then pCargarGrid
    If Button.Index = 6 Then pExportarExcel
    If Button.Index = 7 Then pImprimir
    If Button.Index = 9 Then
        Unload Me
    End If
End Sub

Sub Eliminar()
    On Error GoTo error
    If Fg1.Rows <= 1 Then
        MsgBox "No hay registros", vbExclamation, xTitulo
        Exit Sub
    End If
    If Fg1.Row < 1 Then
        MsgBox "Seleccione correctamente el registro", vbExclamation, xTitulo
        Exit Sub
    End If

    If MsgBox("¿Esta seguro de eliminar el registro?", vbQuestion + vbYesNo, xTitulo) = vbYes Then
        '--eliminar marcacion de asistencia autmatico con origen Dia festivo(feriado)
        pEliminarAsistencia NulosN(Fg1.TextMatrix(Fg1.Row, 1))
        '-----
        xCon.Execute "DELETe * FROM mae_diasfestivos WHERE id = " & NulosN(Fg1.TextMatrix(Fg1.Row, 1)) & "; "
        
        MsgBox "El registro fue eliminado con éxito", vbInformation, xTitulo
        pCargarGrid
    End If
Exit Sub
error:
    SHOW_ERROR Me.Name, "Eliminar"
End Sub

Private Sub Cancelar()
    QueHace = 3
    pHabilitarBotonEditor False
    If Fg1.Rows = 1 Then
        cmd(0).SetFocus
    Else
        Fg1.SetFocus
    End If
    
End Sub

Private Sub CambiarMes()
    mMesActivo = SeleccionaMes(xCon)
    pCargarGrid
    TabOne1.CurrTab = 0
End Sub
Private Sub Modificar()
   '------
    If Fg1.Rows = 1 Then
        MsgBox "No hay Registros", vbExclamation, xTitulo
        Exit Sub
    End If
    If Fg1.Row < 1 Then
        MsgBox "Seleccione correctamente el registro", vbExclamation, xTitulo
        Exit Sub
    End If
    QueHace = 2
    pHabilitarBotonEditor True
    pPonerDatos
    txt(1).SetFocus
End Sub

Private Sub Blanquea()
    LimpiaText txt
    LimpiaText txtfecha
End Sub

Private Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub nuevo()
    QueHace = 1
    Blanquea
    pHabilitarBotonEditor True
    dtpk(0).Value = CDate("00:00:00")
    dtpk(1).Value = CDate("00:00:00")
    txt(1).SetFocus
End Sub


Function Grabar() As Boolean
    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " el Registro", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo salir
    
    Dim RstCab As New ADODB.Recordset
    Dim RstHora As New ADODB.Recordset
    Dim xCod&, xCol&, xFil&
    Dim nSQL As String
    
    On Error GoTo LaCague
    Me.MousePointer = vbHourglass
    xCon.BeginTrans
    
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT top 1 * FROM mae_diasfestivos ", xCon
        xCod = HallaCodigoTabla("mae_diasfestivos", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xCod
    Else
        xCod = NulosN(Fg1.TextMatrix(Fg1.Row, 1))
        
        RST_Busq RstCab, "SELECT * FROM mae_diasfestivos WHERE id =" & xCod & "", xCon
        '--eliminar marcacion de asistencia autmatico con origen Dia festivo(feriado)
        pEliminarAsistencia xCod
        '-----

    End If
    RstCab("anno") = AnoTra
    RstCab("descripcion") = NulosC(txt(1).Text)
    RstCab("fchini") = CDate(txtfecha(0).Valor)
    RstCab("horini") = CDate(dtpk(0).Value)
    RstCab("fchfin") = CDate(txtfecha(1).Valor)
    RstCab("horfin") = CDate(dtpk(1).Value)
    RstCab("observacion") = Trim(txt(2).Text)
    RstCab.Update
    
    '----
    MsgBox "El registro se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    
    xCon.CommitTrans
    Grabar = True
salir:
    Set RstCab = Nothing:    Set RstHora = Nothing
    Me.MousePointer = vbDefault
    Exit Function
LaCague:
    xCon.RollbackTrans
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstHora = Nothing
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
    Grabar = False
End Function


Private Function fValidarDatos() As Boolean
    Dim mRow&, QGrid&, mCodigo&
    Dim band&

    band = validar_Fecha(txtfecha)
    If band <> -1 Then
        MsgBox "Falta ingresar el Campo " & lbltxtfch(band).Caption, vbExclamation, xTitulo
        txtfecha(band).SetFocus
        Exit Function
    End If
    
    band = validar_Fecha(dtpk)
    If band <> -1 Then
        MsgBox "Falta ingresar el Campo " & lbldtpk(band).Caption, vbExclamation, xTitulo
        dtpk(band).SetFocus
        Exit Function
    End If
    '--------------------------------
    If CDate(txtfecha(0).Valor) > CDate(txtfecha(1).Valor) Then
        MsgBox "La fecha de Inicio es superior a la fecha Final" + vbCr + "Modifique los valores para continuar", vbExclamation, xTitulo
        txtfecha(1).SetFocus
        Exit Function
    End If
    
    If (CDate(txtfecha(0).Valor) = CDate(txtfecha(1).Valor)) And (CDate(dtpk(0).Value) > CDate(dtpk(1).Value)) Then
        MsgBox "La Hora de Inicio es superior a la Hora Final" + vbCr + "Modifique los valores para continuar", vbExclamation, xTitulo
        dtpk(1).SetFocus
        Exit Function
    End If
    '--------------------------------
    '--validar que no este registrado
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String

    '--restringir si se programo dias festivos en el intervalo de fechas
    If QueHace = 1 Then
        nSQL = "SELECT id,fchini,fchfin,descripcion FROM mae_diasfestivos " _
                + vbCr + " WHERE " _
                + vbCr + " ((fchini BETWEEN cdate('" & txtfecha(0).Valor & "') AND cdate('" & txtfecha(1).Valor & "')) OR " _
                + vbCr + "  (fchfin BETWEEN cdate('" & txtfecha(0).Valor & "') AND cdate('" & txtfecha(1).Valor & "'))) "
    Else
        nSQL = "SELECT id,fchini,fchfin,descripcion FROM mae_diasfestivos " _
                + vbCr + " WHERE id <> " & NulosN(Fg1.TextMatrix(Fg1.Row, 1)) & " AND " _
                + vbCr + " ((fchini BETWEEN cdate('" & txtfecha(0).Valor & "') AND cdate('" & txtfecha(1).Valor & "')) OR " _
                + vbCr + "  (fchfin BETWEEN cdate('" & txtfecha(0).Valor & "') AND cdate('" & txtfecha(1).Valor & "'))) "
    End If
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.BOF = False Or RstTmp.EOF = False Or RstTmp.RecordCount <> 0 Then
        MsgBox "No se puede " + IIf(QueHace = 1, "agregar", "modificar") + vbCr + _
               "Porque existe un registro del " & Format(RstTmp.Fields("fchini"), "dd/mm/yy") & " al " & Format(RstTmp.Fields("fchfin"), "dd/mm/yy") + vbCr + "Descripción: " + RstTmp.Fields("descripcion"), vbExclamation, xTitulo
        Set RstTmp = Nothing
        Exit Function
    End If
    '--------------------------------
    fValidarDatos = True
End Function
 
Sub Buscar()
    On Error GoTo error
    TabOne1.CurrTab = 0
     
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
   
    Dim xCampos(5, 4) As String
    
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "Descripcion":  xCampos(0, 2) = "1000":  xCampos(0, 3) = "C"
    xCampos(1, 0) = "F.Inicio":     xCampos(1, 1) = "fchini":       xCampos(1, 2) = "1000":  xCampos(1, 3) = "F"
    xCampos(2, 0) = "H.Inicio":     xCampos(2, 1) = "horini":       xCampos(2, 2) = "1000":  xCampos(2, 3) = "F"
    xCampos(3, 0) = "F.Fin":        xCampos(3, 1) = "fchfin":       xCampos(3, 2) = "1000":  xCampos(3, 3) = "F"
    xCampos(4, 0) = "H.Fin":        xCampos(4, 1) = "horfin":       xCampos(4, 2) = "1200":  xCampos(4, 3) = "F"
        
        
    nSQL = "SELECT mae_diasfestivos.* " _
        + vbCr + " FROM mae_diasfestivos WHERE anno = " & AnoTra & " " _
        + vbCr + " ORDER BY mae_diasfestivos.fchini, mae_diasfestivos.horini;"
        
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Dias Festivos", "Descripcion", "Descripcion", Principio
    If xRs.State = 0 Then GoTo salir
    If xRs.EOF = True And xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo salir
    '--
    Dim A&
    Fg1.Row = 1
    For A = 1 To Fg1.Rows - 1
        DoEvents
        Fg1.Row = A
        If NulosN(Fg1.TextMatrix(A, 1)) = xRs("id") Then
            Exit For
        End If
    Next A
    '--
salir:
    Set xRs = Nothing
    Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Buscar"
End Sub

Private Sub pImprimir()
    On Error GoTo error
    Me.MousePointer = vbHourglass
    Dim oPrint As New SGI2_funciones.formularios
    oPrint.Imprimir_x_VSFlexGrid Fg1, "Consulta de Dias Festivos", , lblperiodo(0).Caption, False, True
    Set oPrint = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Set oPrint = Nothing
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"
End Sub

Private Sub pExportarExcel()
    On Error GoTo error
    Dim oExport As New SGI2_funciones.formularios
    oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "Consulta de Dias Festivos", lblperiodo(0).Caption, "", "Dias Festivos"
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportarExcel"
End Sub


Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
End Sub

Private Sub pHabilitarBotonEditor(band As Boolean)
    '--TRUE= MUESTRA LA OPCION PARA SELECCIONAR LA RUTA
    Dim K&
    If band = True Then
        Fg1.Enabled = False
        FraEditor.Top = 1545
        FraEditor.Left = 3285
        LblTituloFrame.Caption = "Agregar Dia Festivo"
    Else
        LblTituloFrame.Caption = "Modificar Dia Festivo"
        Fg1.Enabled = True
    End If
    FraEditor.Visible = band
    habilitar cmd, Not band
    For K = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(K).Enabled = Not band
    Next K
    
End Sub

'****************

Private Sub cb_Click(Index As Integer)
    On Error GoTo error
    Dim obj As New SGI2_funciones.formularios
    obj.HoraSeleccionar dtpk(Index), -1, -1, dtpk(Index).Value
    Set obj = Nothing
    Select Case Index
        Case 0 '--HORA INICIO
            txtfecha(1).SetFocus
        Case 1 '--HORA FIN
            txt(2).SetFocus
    End Select
    Exit Sub
Exit Sub
error:
    
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        Exit Sub
    End If
End Sub

'****************

Private Sub pPonerDatos()
    On Error GoTo error
    Dim mRow&
    
    QueHace = 2
    Agregando = True
    With Fg1
        mRow = .Row
        txt(1).Text = .TextMatrix(mRow, 2)
        If IsDate(.TextMatrix(mRow, 3)) = True Then '--fecha inicio
            txtfecha(0).Valor = CDate(.TextMatrix(mRow, 3))
        Else
            txtfecha(0).Valor = ""
        End If
        If IsDate(.TextMatrix(mRow, 4)) = True Then '--hora inicio
            dtpk(0).Value = CDate(.TextMatrix(mRow, 4))
        Else
            dtpk(0).Value = ""
        End If
        If IsDate(.TextMatrix(mRow, 5)) = True Then '--fecha fin
            txtfecha(1).Valor = CDate(.TextMatrix(mRow, 5))
        Else
            txtfecha(1).Valor = ""
        End If
        If IsDate(.TextMatrix(mRow, 6)) = True Then '--hora fin
            dtpk(1).Value = CDate(.TextMatrix(mRow, 6))
        Else
            dtpk(1).Value = ""
        End If
        txt(2).Text = .TextMatrix(mRow, 7) '--observacion
    End With
    Agregando = False
    Exit Sub
error:
    Agregando = False
    CmdEditor(0).Enabled = False
    SHOW_ERROR Me.Name, "pPonerDatos"
End Sub

Private Sub pConfigurarGrilla()
    With Fg1
        .Rows = 1
        .Cols = 8
        .FixedRows = 1
        .RowHeight(0) = 250
        
        .TextMatrix(0, 1) = "id":           .ColWidth(1) = 0:    .ColAlignment(1) = flexAlignLeftCenter
        
        .TextMatrix(0, 2) = "Descripción":  .ColWidth(2) = 4000: .ColAlignment(2) = flexAlignLeftCenter:     .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 3) = "F.Inicio":     .ColWidth(3) = 900:  .ColAlignment(3) = flexAlignCenterCenter:   .Row = 0: .Col = 3: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 4) = "H.Inicio":     .ColWidth(4) = 1150: .ColAlignment(4) = flexAlignCenterCenter:   .Row = 0: .Col = 4: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 5) = "F.Fin":        .ColWidth(5) = 900:  .ColAlignment(5) = flexAlignCenterCenter:   .Row = 0: .Col = 5: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 6) = "H.Fin":        .ColWidth(6) = 1150: .ColAlignment(6) = flexAlignCenterCenter:   .Row = 0: .Col = 6: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 7) = "Observación":  .ColWidth(7) = 3200: .ColAlignment(7) = flexAlignLeftCenter:     .Row = 0: .Col = 7: .CellAlignment = flexAlignLeftCenter
        
        .ColFormat(3) = FORMAT_DATE
        .ColFormat(4) = FORMAT_HORA_AL_SEGUNDO
        .ColFormat(5) = FORMAT_DATE
        .ColFormat(6) = FORMAT_HORA_AL_SEGUNDO
        .SelectionMode = flexSelectionByRow
    End With
    '*****************************************
    DoEvents
End Sub

Private Sub pEliminarAsistencia(mIdCodigo&)
    '--Eliminando los registros de asistencia
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim mIdEmp&
    '--buscando se hay dias que se registraron las asistencias
    '--idori=7:dia festivo segun tabla pla_origenes
    nSQL = "SELECT pla_licencia.idemp, pla_marcacion.dia, pla_marcaciondet.idori, pla_marcaciondet.idmarca " _
        + vbCr + " FROM pla_licencia, pla_marcacion INNER JOIN pla_marcaciondet ON pla_marcacion.id = pla_marcaciondet.idmarca " _
        + vbCr + " WHERE (((pla_marcacion.dia) Between [pla_licencia].[fchini] And [pla_licencia].[fchfin]) " & _
                  "AND ((pla_marcaciondet.idori)=7) AND ((pla_licencia.id)=" & mIdCodigo & "));"
                  
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.RecordCount <> 0 Then
        RstTmp.MoveFirst
        nSQL = ""
        mIdEmp = RstTmp.Fields("idemp")
        Do While Not RstTmp.EOF
            nSQL = nSQL & RstTmp.Fields("idmarca") & ","
            RstTmp.MoveNext
        Loop
        If nSQL <> "" Then nSQL = " (" + Left(nSQL, Len(nSQL) - 1) + ") "
        '--marcacion
        xCon.Execute "DELETE FROM pla_marcaciondet " & _
                     "WHERE idemp = " & mIdEmp & " AND idori=7 AND idmarca In " & nSQL & " ;"
        '--tipos de horas
        xCon.Execute "DELETE FROM pla_marcacionhora " & _
                     "WHERE idemp = " & mIdEmp & " AND idhora =10 AND idmarca In " & nSQL & " ;"
        '--10 hora feriado
    End If
    Set RstTmp = Nothing

End Sub

Private Sub pic_Click()
    CmdEditor_Click 1
End Sub



