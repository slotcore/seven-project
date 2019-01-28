VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVacaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Vacaciones"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   11685
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      Height          =   540
      Index           =   1
      Left            =   30
      TabIndex        =   21
      Top             =   6255
      Width           =   11625
      Begin VB.CommandButton cmd 
         Caption         =   "Modificar"
         Height          =   345
         Index           =   1
         Left            =   1320
         TabIndex        =   24
         Top             =   105
         Width           =   1200
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Agregar"
         Height          =   345
         Index           =   0
         Left            =   75
         TabIndex        =   23
         Top             =   105
         Width           =   1200
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Eliminar"
         Height          =   345
         Index           =   2
         Left            =   2880
         TabIndex        =   22
         Top             =   105
         Width           =   1200
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   3
         X1              =   11610
         X2              =   11610
         Y1              =   -30
         Y2              =   5455
      End
      Begin VB.Line lin 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   3
         X1              =   -30
         X2              =   12000
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   2
         X1              =   -15
         X2              =   12000
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      Height          =   420
      Index           =   12
      Left            =   30
      TabIndex        =   19
      Top             =   375
      Width           =   11625
      Begin VB.Label lblperiodo 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   1770
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   0
         X1              =   -15
         X2              =   12000
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line lin 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   -30
         X2              =   12000
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   11610
         X2              =   11610
         Y1              =   -30
         Y2              =   5455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   4
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   380
      End
   End
   Begin VB.Frame FraEditor 
      BorderStyle     =   0  'None
      Height          =   4320
      Left            =   3435
      TabIndex        =   9
      Top             =   1410
      Visible         =   0   'False
      Width           =   5340
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   5055
         Picture         =   "FrmVacaciones.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   26
         ToolTipText     =   "Cerrar"
         Top             =   75
         Width           =   195
      End
      Begin VB.Frame Frame3 
         Caption         =   "[ De los Dias ]"
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
         Height          =   2190
         Left            =   120
         TabIndex        =   18
         Top             =   1500
         Width           =   5100
         Begin VB.CommandButton CmdDet 
            Caption         =   "&Agregar"
            Height          =   330
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   1785
            Width           =   1395
         End
         Begin VB.CommandButton CmdDet 
            Caption         =   "&Eliminar"
            Height          =   330
            Index           =   1
            Left            =   1620
            TabIndex        =   4
            Top             =   1785
            Width           =   1395
         End
         Begin VSFlex7Ctl.VSFlexGrid fg2 
            Height          =   1515
            Left            =   105
            TabIndex        =   5
            Top             =   225
            Width           =   4875
            _cx             =   8599
            _cy             =   2672
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmVacaciones.frx":02EC
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
            Ellipsis        =   1
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
      Begin VB.Frame Frame2 
         Caption         =   "[ Del Pago ]"
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
         Height          =   705
         Left            =   120
         TabIndex        =   15
         Top             =   765
         Width           =   5100
         Begin VB.ComboBox cbMes 
            Height          =   315
            Left            =   2565
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   285
            Width           =   2055
         End
         Begin VB.TextBox txt 
            Height          =   315
            Index           =   1
            Left            =   570
            MaxLength       =   4
            TabIndex        =   1
            Text            =   "txt(1)"
            Top             =   285
            Width           =   780
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mes"
            Height          =   195
            Index           =   3
            Left            =   2160
            TabIndex        =   17
            Top             =   375
            Width           =   300
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Año"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   16
            Top             =   375
            Width           =   285
         End
      End
      Begin VB.CommandButton cb 
         Height          =   225
         Index           =   0
         Left            =   1380
         Picture         =   "FrmVacaciones.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Seleccione el Personal"
         Top             =   450
         Width           =   210
      End
      Begin VB.CommandButton CmdEditor 
         Caption         =   "&Grabar"
         Height          =   420
         Index           =   0
         Left            =   1515
         TabIndex        =   6
         Top             =   3810
         Width           =   1020
      End
      Begin VB.CommandButton CmdEditor 
         Caption         =   "&Cancelar"
         Height          =   420
         Index           =   1
         Left            =   2715
         TabIndex        =   7
         Top             =   3810
         Width           =   1020
      End
      Begin VB.TextBox txt_cb 
         Height          =   300
         Index           =   0
         Left            =   855
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "txt_cb(0)"
         Top             =   420
         Width           =   765
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
         Left            =   4050
         TabIndex        =   13
         Top             =   420
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Personal"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   510
         Width           =   615
      End
      Begin VB.Label LblTituloFrame 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Editor de Vacaciones"
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
         TabIndex        =   10
         Top             =   90
         Width           =   1830
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5325
         X2              =   5325
         Y1              =   -135
         Y2              =   4755
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
         Index           =   0
         X1              =   -30
         X2              =   5640
         Y1              =   4290
         Y2              =   4305
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
         X2              =   5235
         Y1              =   3750
         Y2              =   3750
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   300
         Index           =   1
         Left            =   30
         Top             =   45
         Width           =   5250
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
         Left            =   1620
         TabIndex        =   14
         Top             =   420
         Width           =   3615
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
            Picture         =   "FrmVacaciones.frx":04BC
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVacaciones.frx":0A00
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVacaciones.frx":0D92
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVacaciones.frx":0F16
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVacaciones.frx":136A
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVacaciones.frx":1482
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVacaciones.frx":19C6
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVacaciones.frx":1F0A
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVacaciones.frx":201E
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVacaciones.frx":2132
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVacaciones.frx":2586
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVacaciones.frx":26F2
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVacaciones.frx":2C3A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   1005
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
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   5385
      Left            =   30
      TabIndex        =   25
      Top             =   810
      Width           =   11625
      _cx             =   20505
      _cy             =   9499
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
      FormatString    =   $"FrmVacaciones.frx":2FCC
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
   Begin VB.Menu Menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu Menu2_1 
         Caption         =   "&Agregar"
      End
      Begin VB.Menu Menu2_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu2_3 
         Caption         =   "&Eliminar"
      End
   End
End
Attribute VB_Name = "FrmVacaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim Agregando As Boolean

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0 '--Add Vacaciones
            nuevo
        Case 1 '--Modificar Vacaciones
            Modificar
        Case 2 '--Eliminar Vacaciones
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

Private Sub pCargarGrid()
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL  As String
    On Error GoTo error
    lblperiodo(0).Caption = "Periodo: " & AnoTra
    nSQL = "SELECT pla_vacaciones.*, pla_empleados.numdoc, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres " _
        + vbCr + " FROM pla_empleados RIGHT JOIN pla_vacaciones ON pla_empleados.id = pla_vacaciones.idemp " _
        + vbCr + " WHERE (((pla_vacaciones.anno) = " & AnoTra & ")) " _
        + vbCr + " ORDER BY pla_vacaciones.mespago;"


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
                    .TextMatrix(.Rows - 1, 2) = NulosC(RstTmp.Fields("idemp"))
                    .TextMatrix(.Rows - 1, 3) = NulosC(RstTmp.Fields("numdoc"))
                    .TextMatrix(.Rows - 1, 4) = NulosC(RstTmp.Fields("nombres"))
                    .TextMatrix(.Rows - 1, 5) = NulosC(RstTmp.Fields("annopago"))
                    .TextMatrix(.Rows - 1, 6) = NomMes(RstTmp.Fields("mespago"))
                    .TextMatrix(.Rows - 1, 7) = NulosN(RstTmp.Fields("mespago"))
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
    ElseIf KeyCode = 46 Then
        Eliminar
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Fg1.Enabled = False Then Exit Sub
    If Button = 2 Then
        PopupMenu menu1
    End If
End Sub

Private Sub fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Or Col = 3 Then
        '--invocar al formulario de fecha
        Dim obj As New SGI2_funciones.formularios
        obj.FechaSeleccionar fg2, Row, Col, fg2.TextMatrix(Row, Col)
        Set obj = Nothing
    End If
    Exit Sub
salir:
    Agregando = False
End Sub

Private Sub fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    If Agregando = True Then Exit Sub
    If IsDate(fg2.TextMatrix(Row, Col)) = False Then
        fg2.TextMatrix(Row, Col) = ""
        fg2.TextMatrix(Row, 4) = ""
        Exit Sub
    End If
    Select Case Col
        Case 2
            If IsDate(fg2.TextMatrix(Row, 3)) = True Then
                If CDate(fg2.TextMatrix(Row, 3)) < CDate(fg2.TextMatrix(Row, 2)) Then
                    MsgBox "La Fecha Inicial es Superior a la Fecha Final", vbExclamation, xTitulo
                    fg2.TextMatrix(Row, Col) = ""
                End If
            End If
        Case 3
            If IsDate(fg2.TextMatrix(Row, 2)) = True Then
                If CDate(fg2.TextMatrix(Row, 3)) < CDate(fg2.TextMatrix(Row, 2)) Then
                    MsgBox "La Fecha Final es Inferior a la Fecha Inicial", vbExclamation, xTitulo
                    fg2.TextMatrix(Row, Col) = ""
                End If
            End If
    End Select
    If IsDate(fg2.TextMatrix(Row, 2)) = True And IsDate(fg2.TextMatrix(Row, 3)) = True Then
        fg2.TextMatrix(Row, 4) = DateDiff("d", fg2.TextMatrix(Row, 2), fg2.TextMatrix(Row, 3)) + 1
    Else
        fg2.TextMatrix(Row, 4) = ""
    End If
    
    Exit Sub
    
error:
    SHOW_ERROR Me.Name, "fg2_CellChanged"
End Sub


Private Sub fg2_EnterCell()
    If QueHace = 3 Then Exit Sub
    If fg2.Col = 4 Then
        fg2.Editable = flexEDNone
    Else
        fg2.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub fg2_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 45 Then
        CmdDet_Click 0 '--agregar
    ElseIf KeyCode = 46 Then
        CmdDet_Click 1 '--eliminar
    End If
End Sub

Private Sub fg2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu Menu2
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
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
    
    '--
    '--
    Llenar_Mes cbMes
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

Private Sub Menu2_1_Click()
    CmdDet_Click 0
End Sub

Private Sub Menu2_3_Click()
    CmdDet_Click 1
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
        '--eliminar asistencia  ---------------------------------
        pEliminarAsistencia NulosN(Fg1.TextMatrix(Fg1.Row, 1))
        '---------------------------------------------------------
        xCon.Execute "DELETe * FROM pla_vacacionesdet WHERE idvac = " & NulosN(Fg1.TextMatrix(Fg1.Row, 1)) & "; "
        xCon.Execute "DELETe * FROM pla_vacaciones WHERE id = " & NulosN(Fg1.TextMatrix(Fg1.Row, 1)) & "; "
        
        pCargarGrid
        MsgBox "El registro fue eliminado con éxito", vbInformation, xTitulo
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
    xMes = SeleccionaMes(xCon)
    pCargarGrid
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
    pHabilitarBotonEditor True
    pPonerDatos
    QueHace = 2
    LblTituloFrame.Caption = "Modificar Vacaciones"
    txt(1).SetFocus
End Sub

Private Sub Blanquea()
    LimpiaText txt
    LimpiaText txt_cb
    LimpiaText lbl_cod
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
    cbMes.ListIndex = -1
    LblTituloFrame.Caption = "Agregar Vacaciones"
    txt_cb(0).SetFocus
End Sub


Function Grabar() As Boolean
    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "grabar", "Modificar") + " el Registro", vbQuestion + vbYesNo, xTitulo) = vbNo Then GoTo salir
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim RstHora As New ADODB.Recordset
    Dim RstTmp As New ADODB.Recordset
    Dim xCod&, xCol&, xFil&
    Dim nSQL As String
    
    On Error GoTo LaCague
    Me.MousePointer = vbHourglass
    xCon.BeginTrans
    
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT top 1 * FROM pla_vacaciones ", xCon
        xCod = HallaCodigoTabla("pla_vacaciones", xCon, "id")
        RstCab.AddNew
        RstCab("id") = xCod
    Else
        xCod = NulosN(Fg1.TextMatrix(Fg1.Row, 1))
        
        RST_Busq RstCab, "SELECT * FROM pla_vacaciones WHERE id =" & xCod & "", xCon
        '--eliminando el detalle de las vacaciones
        xCon.Execute "DELETE FROM pla_vacacionesdet WHERE idvac =" & xCod & ""
        '******************************************************************************************************
        '--Eliminando los registros de asistencia
        pEliminarAsistencia xCod
        '******************************************************************************************************

    End If
    RST_Busq RstDet, "SELECT top 1 * FROM pla_vacacionesdet ", xCon
    '-----------
    RstCab("anno") = AnoTra
    RstCab("idemp") = NulosN(lbl_cod(0).Caption)
    RstCab("annopago") = NulosN(txt(1).Text)
    RstCab("mespago") = cbMes.ListIndex + 1
    RstCab.Update
    '-------
    Dim dFecha As Date
    '----del las fechas de las vacaciones
    With fg2
        For xFil = 1 To .Rows - 1
            RstDet.AddNew
            RstDet("idvac") = xCod
            RstDet("corr") = xFil
            RstDet("fchini") = CDate(.TextMatrix(xFil, 2))
            RstDet("fchfin") = CDate(.TextMatrix(xFil, 3))
            RstDet("numdias") = DateDiff("d", CDate(.TextMatrix(xFil, 2)), CDate(.TextMatrix(xFil, 3))) + 1
            RstDet.Update
            '******************************************************************************************************
            '--generar los registros de asistencia en automatico
            '----
            For dFecha = CDate(.TextMatrix(xFil, 2)) To CDate(.TextMatrix(xFil, 3))
                pMacacionDia dFecha, e_Asist_Vacaciones, NulosN(lbl_cod(0).Caption)
            Next dFecha
            '******************************************************************************************************
            
        Next xFil
    End With
    '-------
    '--falta registro en asistencia por vacaciones
    xCon.CommitTrans
    

    MsgBox "El registro se " + IIf(QueHace = 1, "grabó", "modificó") + " con éxito", vbInformation, xTitulo
    
    Grabar = True
salir:
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstHora = Nothing:    Set RstTmp = Nothing
    Me.MousePointer = vbDefault
    Exit Function
LaCague:
    xCon.RollbackTrans
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:    Set RstDet = Nothing:    Set RstHora = Nothing:    Set RstTmp = Nothing
    SHOW_ERROR Me.Name, "Grabar", True, "No se pudo guardar el registro por el siguiente motivo :"
    Grabar = False
End Function


Private Function fValidarDatos() As Boolean
    
    Dim mRow&, QGrid&, mCodigo&
    Dim band&

    '--------------------------------
    'ver si tiene horario no este registrado el horario
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    
    nSQL = "SELECT mae_horariohora.idhora, mae_tipohora.descripcion, mae_horariohora.hingreso, mae_horariohora.hsalida " _
        + vbCr + " FROM mae_tipohora INNER JOIN (mae_horariohora INNER JOIN mae_horarioemp ON mae_horariohora.idhor = mae_horarioemp.idhor) ON mae_tipohora.id = mae_horariohora.idhora " _
        + vbCr + " Where (((mae_horariohora.idhora) = 1) And ((mae_horarioemp.IdEmp) = " & NulosN(txt_cb(0).Text) & ") And ((mae_horarioemp.vigencia) = -1)) " _
        + vbCr + " ORDER BY mae_tipohora.prioridad;"
    RST_Busq RstTmp, nSQL, xCon
    If RstTmp.RecordCount = 0 Then
        MsgBox "El personal no tiene Horario" + vbCr + "Configure el horario a " & lbl_cb(0).Caption, vbExclamation, xTitulo
        Set RstTmp = Nothing
        Exit Function
    End If
    Set RstTmp = Nothing
    '--------------------------------
    If NulosN(lbl_cod(0).Caption) = 0 Then
        MsgBox "Falta ingresar el Personal", vbExclamation, xTitulo
        txt_cb(0).SetFocus
        Exit Function
    End If
    
    If NulosN(txt(1).Text) = 0 Then
        MsgBox "Falta ingresar el Año de Pago", vbExclamation, xTitulo
        txt(1).SetFocus
        Exit Function
    End If
    If cbMes.ListIndex = -1 Then
        MsgBox "Falta ingresar el mes de Pago", vbExclamation, xTitulo
        cbMes.SetFocus
        SendKeys "{F4}"
        Exit Function
    End If
    '--------------------------------
    With fg2
        For mRow = 1 To .Rows - 1
            If IsDate(.TextMatrix(mRow, 2)) = False Then
                MsgBox "Ingrese la Fecha de Inicio", vbExclamation, xTitulo
                Agregando = True:  .Row = mRow:  .Col = 2: Agregando = False
                fg2.SetFocus
                Exit Function
            ElseIf IsDate(.TextMatrix(mRow, 3)) = False Then
                MsgBox "Ingrese la Fecha Final", vbExclamation, xTitulo
                Agregando = True:  .Row = mRow:  .Col = 3: Agregando = False
                fg2.SetFocus
                Exit Function
            End If
        Next mRow
    End With
    fValidarDatos = True
End Function
 


Sub Buscar()
    On Error GoTo error
     
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
   
    Dim xCampos(5, 4) As String
    
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "Descripcion":  xCampos(0, 2) = "1000":  xCampos(0, 3) = "C"
    xCampos(1, 0) = "F.Inicio":     xCampos(1, 1) = "fchini":       xCampos(1, 2) = "1000":  xCampos(1, 3) = "F"
    xCampos(2, 0) = "H.Inicio":     xCampos(2, 1) = "horini":       xCampos(2, 2) = "1000":  xCampos(2, 3) = "F"
    xCampos(3, 0) = "F.Fin":        xCampos(3, 1) = "fchfin":       xCampos(3, 2) = "1000":  xCampos(3, 3) = "F"
    xCampos(4, 0) = "H.Fin":        xCampos(4, 1) = "horfin":       xCampos(4, 2) = "1200":  xCampos(4, 3) = "F"
        
        
    nSQL = "SELECT pla_vacaciones.* " _
        + vbCr + " FROM pla_vacaciones WHERE anno = " & AnoTra & " " _
        + vbCr + " ORDER BY pla_vacaciones.fchini, pla_vacaciones.horini;"
        
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
    oPrint.Imprimir_x_VSFlexGrid Fg1, "Consulta de Vacaciones", , lblperiodo(0).Caption, False, True
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
    oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "Consulta de Vacaciones", lblperiodo(0).Caption, "", "Vacaciones"
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
        fg2.Rows = 1
    Else
        Fg1.Enabled = True
    End If
    GRID_COMBOLIST fg2, 2
    GRID_COMBOLIST fg2, 3
    FraEditor.Visible = band
    habilitar cmd, Not band
    For K = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(K).Enabled = Not band
    Next K
    
End Sub

'****************

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
    
    Agregando = True
    With Fg1
        mRow = .Row
        '--personal
        txt_cb(0).Text = NulosN(.TextMatrix(mRow, 2))
        lbl_cb(0).Caption = NulosC(.TextMatrix(mRow, 4))
        lbl_cod(0).Caption = NulosN(.TextMatrix(mRow, 2))
        
        txt(1).Text = NulosN(.TextMatrix(mRow, 5))
        
        If NulosN(.TextMatrix(mRow, 7)) <> 0 Then
            cbMes.ListIndex = NulosN(.TextMatrix(mRow, 7)) - 1
        Else
            cbMes.ListIndex = -1
        End If
    End With
    '*****
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    
    nSQL = "SELECT pla_vacacionesdet.* From pla_vacacionesdet " _
        + vbCr + " WHERE (((pla_vacacionesdet.idvac) = " & NulosN(Fg1.TextMatrix(mRow, 1)) & ")) " _
        + vbCr + " ORDER BY pla_vacacionesdet.fchini; "

    Me.MousePointer = vbHourglass
    RST_Busq RstTmp, nSQL, xCon
    '---------------
    If RstTmp.RecordCount <> 0 Then
        Agregando = True
        If RstTmp.RecordCount <> 0 Then
            fg2.Rows = 1
            RstTmp.MoveFirst
            Do While Not RstTmp.EOF
                With fg2
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = .Row ' RstTmp.Bookmark
                    .TextMatrix(.Rows - 1, 2) = NulosC(RstTmp.Fields("fchini"))
                    .TextMatrix(.Rows - 1, 3) = NulosC(RstTmp.Fields("fchfin"))
                    .TextMatrix(.Rows - 1, 4) = NulosC(RstTmp.Fields("numdias")) '
                    RstTmp.MoveNext
                End With
            Loop
        End If
    End If
    Set RstTmp = Nothing
    '---------------
    Me.MousePointer = vbDefault
    '*****
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
        
        .TextMatrix(0, 1) = "id":                   .ColWidth(1) = 0:   .ColAlignment(1) = flexAlignLeftCenter
        .TextMatrix(0, 2) = "idEmp":                .ColWidth(2) = 0:   .ColAlignment(2) = flexAlignLeftCenter
        .TextMatrix(0, 3) = "D.N.I.":               .ColWidth(3) = 1000:  .ColAlignment(3) = flexAlignCenterCenter:     .Row = 0: .Col = 3: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 4) = "Apellidos y Nombres":  .ColWidth(4) = 4500:  .ColAlignment(4) = flexAlignLeftCenter:     .Row = 0: .Col = 4: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 5) = "Año Pago":             .ColWidth(5) = 1200:  .ColAlignment(5) = flexAlignCenterCenter:   .Row = 0: .Col = 5: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 6) = "Mes Pago":             .ColWidth(6) = 1500:  .ColAlignment(6) = flexAlignLeftCenter:     .Row = 0: .Col = 6: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(0, 7) = "Mes_Pago":             .ColWidth(7) = 0:
        
        .SelectionMode = flexSelectionByRow
    End With
    '*****************************************
    With fg2
        .ColWidth(1) = 0:
        .ColEditMask(2) = "##/##/####"
        .ColEditMask(3) = "##/##/####"
    End With
    '*****************************************
    DoEvents
End Sub


Private Sub cb_Click(Index As Integer)
    If QueHace = 3 Then Exit Sub
    On Error GoTo error
    Dim xRs As New ADODB.Recordset
    pBuscarPersonal xRs, True
    If xRs.State = 1 Then
        txt_cb(0) = xRs.Fields("id") & "" '--TEXTO A MOSTRAR
        lbl_cb(0).Caption = xRs.Fields("nombres") & "" '--NOMBRE
        lbl_cod(0).Caption = xRs.Fields("id") & "" '--CODIGO
        lbl_cb(0).ToolTipText = xRs.Fields("nombres") & "" '--NOMBRE
        txt_cb(0).SetFocus
    End If
    Set xRs = Nothing
    txt(1).SetFocus
Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "cb_Click(" + CStr(Index) + ")"
End Sub

Private Sub txt_cb_Change(Index As Integer)
    If QueHace = 3 Then Exit Sub
    If txt_cb(Index).Text = "" Then
        Me.lbl_cb(Index).Caption = ""
        Me.lbl_cod(Index).Caption = ""
    End If
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
        SendKeys vbTab
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
        Case 0 '--personal
            nSQL = "SELECT pla_empleados.id, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombre, pla_empleados.id as cod,pla_empleados.numdoc, mae_dociden.abrev AS tipodoc, mae_sexo.abrev AS sexo, Format([pla_empleados].[fchnac],'dd/mm/yyyy') AS fchnac, pla_empleados.numtel, pla_empleados.email " _
                + vbCr + " FROM mae_sexo RIGHT JOIN (mae_dociden RIGHT JOIN pla_empleados ON mae_dociden.id = pla_empleados.idtipdoc) ON mae_sexo.id = pla_empleados.idsex " _
                + vbCr + " WHERE pla_empleados.id = " & NulosN(txt_cb(Index).Text) & " " _
                + vbCr + " AND pla_empleados.id Not In (SELECT idemp FROM pla_vacaciones WHERE anno = " & AnoTra & " )"
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
        txt_cb(Index).Text = ""
    End If
    '--------------
    Select Case Index
        Case 0 '--personal
            If Agregando = False Then txt(1).SetFocus
    End Select
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


'****************

Private Sub pRegistroAdd()
    Dim mCol%
    If QueHace = 3 Then Exit Sub
    Agregando = True
    If fg2.Rows > 1 Then
        If mCol = 0 Then
            If IsDate(fg2.TextMatrix(fg2.Rows - 1, 2)) = False Then
                MsgBox "Falta ingresar la Fecha de Inicio", vbExclamation, xTitulo
                mCol = 2
            ElseIf IsDate(fg2.TextMatrix(fg2.Rows - 1, 3)) = False Then
                MsgBox "Falta ingresar la Fecha Final", vbExclamation, xTitulo
                mCol = 3
            Else
                fg2.AddItem ""
                mCol = 2
            End If
        End If
    Else
        fg2.AddItem ""
        mCol = 2
    End If
    fg2.Row = fg2.Rows - 1
    fg2.Col = mCol
    fg2.SetFocus
    Agregando = False
End Sub

Private Sub pRegistroDel()
    If fg2.Rows = 1 Then Exit Sub
    If fg2.Row < 1 Then Exit Sub
    If MsgBox("Seguro desea eliminar el Registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    fg2.RemoveItem fg2.Row
End Sub

Private Sub CmdDet_Click(Index As Integer)
    Select Case Index
        Case 0 '--Agregar
            pRegistroAdd
        Case 1 '--Eliminar
            pRegistroDel
    End Select
End Sub

'********************


Private Sub pEliminarAsistencia(mIdCodigo&)
    '--Eliminando los registros de asistencia
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    Dim mIdEmp&
    '--buscando se hay dias que se registraron las asistencias
     '--idori=2:permiso segun tabla pla_origenes
    nSQL = "SELECT pla_marcaciondet.idemp,pla_marcacion.dia, pla_vacacionesdet.fchini, pla_vacacionesdet.fchfin, pla_marcaciondet.idmarca " _
        + vbCr + " FROM pla_vacaciones INNER JOIN pla_vacacionesdet ON pla_vacaciones.id = pla_vacacionesdet.idvac, pla_marcacion INNER JOIN pla_marcaciondet ON pla_marcacion.id = pla_marcaciondet.idmarca " _
        + vbCr + " WHERE (((pla_vacaciones.id)= " & mIdCodigo & " ) AND ((pla_marcacion.dia) Between [pla_vacacionesdet].[fchini] And [pla_vacacionesdet].[fchfin]) AND ((pla_marcaciondet.idori)=4)); "

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
                     "WHERE idemp = " & mIdEmp & " AND idori=4 AND idmarca In " & nSQL & " ;"
            '--obs::idori=4,segun tabla pla_origenes
        '--tipos de horas
        xCon.Execute "DELETE FROM pla_marcacionhora " & _
                     "WHERE idemp = " & mIdEmp & " AND idhora =4 AND idmarca In " & nSQL & " ;"
            '--obs::4:hora vacaciones
        
        
    End If
    Set RstTmp = Nothing

End Sub

Private Sub pic_Click()
    CmdEditor_Click 1
End Sub

