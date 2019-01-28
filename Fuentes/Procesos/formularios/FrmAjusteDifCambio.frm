VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmAjusteDifCambio 
   Caption         =   "Herramientas - Ajuste po Diferencia de Cambio"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   7830
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   2115
      Begin VB.OptionButton OptEstado 
         Caption         =   "Pendiente"
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1425
      End
      Begin VB.OptionButton OptEstado 
         Caption         =   "Cancelado"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   120
         Value           =   -1  'True
         Width           =   1395
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "[ Hasta el dia ]"
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
      Height          =   600
      Left            =   5730
      TabIndex        =   14
      Top             =   360
      Width           =   2025
      Begin AspaTextBoxFecha.TextBoxFecha TxtFecha 
         Height          =   315
         Left            =   630
         TabIndex        =   15
         Top             =   270
         Width           =   1350
         _ExtentX        =   2381
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
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   315
         Width           =   450
      End
   End
   Begin VB.OptionButton OptSel 
      Caption         =   "Ajuste de Bancos"
      Height          =   225
      Index           =   1
      Left            =   9990
      TabIndex        =   13
      Top             =   720
      Width           =   1665
   End
   Begin VB.OptionButton OptSel 
      Caption         =   "Otros Ajuste"
      Height          =   195
      Index           =   0
      Left            =   9990
      TabIndex        =   12
      Top             =   450
      Value           =   -1  'True
      Width           =   1725
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Eliminar"
      Height          =   435
      Left            =   10170
      TabIndex        =   11
      Top             =   2730
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccionar "
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
      Height          =   600
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   5685
      Begin VB.CommandButton CmdBusProv 
         Height          =   230
         Left            =   5340
         Picture         =   "FrmAjusteDifCambio.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   270
         Width           =   210
      End
      Begin VB.TextBox TxtAjuste 
         Height          =   300
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "TxtAjuste"
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label LblIdLibro 
         AutoSize        =   -1  'True
         Caption         =   "LblIdLibro"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3690
         TabIndex        =   20
         Top             =   90
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label LblIdAjuste 
         AutoSize        =   -1  'True
         Caption         =   "LblIdAjuste"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2070
         TabIndex        =   10
         Top             =   90
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label LblIdMon 
         AutoSize        =   -1  'True
         Caption         =   "LblIdMon"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4650
         TabIndex        =   9
         Top             =   90
         Visible         =   0   'False
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   735
      Left            =   3180
      TabIndex        =   1
      Top             =   2580
      Visible         =   0   'False
      Width           =   5805
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   285
         Left            =   90
         TabIndex        =   2
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   5820
         Y1              =   15
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   15
         Y1              =   30
         Y2              =   1170
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   1
         X1              =   5790
         X2              =   5790
         Y1              =   15
         Y2              =   1155
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   5820
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Corregiendo Asientos"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   135
         Width           =   1500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1695
         TabIndex        =   3
         Top             =   150
         Width           =   45
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid fg1 
      Height          =   4620
      Left            =   30
      TabIndex        =   0
      Top             =   1950
      Width           =   11670
      _cx             =   20585
      _cy             =   8149
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
      BackColor       =   14745342
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   128
      ForeColorSel    =   -2147483634
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmAjusteDifCambio.frx":0132
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9150
      Top             =   -240
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
            Picture         =   "FrmAjusteDifCambio.frx":0207
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambio.frx":074B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambio.frx":0ADD
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambio.frx":0C61
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambio.frx":10B5
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambio.frx":11CD
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambio.frx":1711
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambio.frx":1C55
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambio.frx":1D69
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambio.frx":1E7D
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambio.frx":22D1
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambio.frx":243D
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAjusteDifCambio.frx":2985
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame6 
      Height          =   915
      Left            =   0
      TabIndex        =   21
      Top             =   990
      Width           =   11655
      Begin VB.CommandButton CmdBusGan 
         Height          =   240
         Left            =   2160
         Picture         =   "FrmAjusteDifCambio.frx":2D17
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   210
         Width           =   240
      End
      Begin VB.CommandButton CmdBusPer 
         Height          =   240
         Left            =   2160
         Picture         =   "FrmAjusteDifCambio.frx":2E49
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   570
         Width           =   240
      End
      Begin VB.TextBox TxtPerdida 
         Height          =   300
         Left            =   1155
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   24
         Text            =   "TxtPerdida"
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox TxtGanancia 
         Height          =   300
         Left            =   1155
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   25
         Text            =   "TxtGanancia"
         Top             =   180
         Width           =   1275
      End
      Begin VB.Line Line1 
         X1              =   10350
         X2              =   10350
         Y1              =   210
         Y2              =   840
      End
      Begin VB.Label lblTotReg 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblTotReg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10800
         TabIndex        =   37
         Top             =   420
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tot. Reg:"
         Height          =   195
         Left            =   10740
         TabIndex        =   36
         Top             =   180
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Ganáncia:"
         Height          =   195
         Left            =   6900
         TabIndex        =   35
         Top             =   270
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Pérdida:"
         Height          =   195
         Index           =   0
         Left            =   6900
         TabIndex        =   34
         Top             =   600
         Width           =   990
      End
      Begin VB.Label lblGan 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblGan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8250
         TabIndex        =   33
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label lblPer 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblPer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8250
         TabIndex        =   32
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label LblIdCtaPer 
         AutoSize        =   -1  'True
         Caption         =   "LblIdCtaPer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4950
         TabIndex        =   27
         Top             =   630
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label LblIdCtaGan 
         AutoSize        =   -1  'True
         Caption         =   "LblIdCtaGan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4950
         TabIndex        =   26
         Top             =   270
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cta Ganáncia"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   31
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cta Pérdida"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   825
      End
      Begin VB.Label LblDescGanancia 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblDescGanancia"
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
         Left            =   2430
         TabIndex        =   29
         Top             =   180
         Width           =   4035
      End
      Begin VB.Label LblDescPerdida 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblDescPerdida"
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
         Left            =   2430
         TabIndex        =   28
         Top             =   540
         Width           =   4035
      End
   End
End
Attribute VB_Name = "FrmAjusteDifCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim SeEjecuto As Boolean

Dim SGI_JC As New SGI2_funciones.JC_Varios
Dim SGI_JC1 As New SGI2_funciones.JC_VSFlexGrid

Dim RstFrm As New ADODB.Recordset
Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim NumRegistro As String

Dim mIdTipPer As Integer '--variable para identificar el proveedor, cliente, banco en diario

Private Sub Command1_Click()
''    xCon.Execute "delete from con_diario where idlib = 6 and ajuste = 1 "
''    MsgBox "Se eliminaron los ajustes automaticos", vbInformation, xTitulo

End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    lblGan.Caption = "0.00"
    lblPer.Caption = "0.00"
    lblTotReg.Caption = 0
    
    TxtAjuste.Text = ""
    
    TxtGanancia.Text = ""
    TxtPerdida.Text = ""
    
    
    
    Configurar_Grilla
    SeEjecuto = True
    TxtFecha.Valor = Date
    TxtAjuste.SetFocus
    
End Sub

Private Sub Form_Deactivate()
    
'    On Error Resume Next
'
'    Err.Clear
'    Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then Unload Me

End Sub

Private Sub Form_Load()
    SeEjecuto = False
    SGI_JC.CentrarFrm Me
    
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
  
    If Me.Height > 3000 Then
        
        '--detalle
        Fg1.Top = 1950
        Fg1.Width = Me.Width - 200
        Fg1.Height = Me.Height - 2400
    End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SGI_JC = Nothing
    Set SGI_JC1 = Nothing
    Set RstFrm = Nothing
    
End Sub


Private Sub Configurar_Grilla()

    With Fg1
        '-----
        .Rows = 2
        .FixedRows = 2
        .Cols = 30
        
        .ColWidth(0) = 200
        '--DATOS DE FILA
        
        SGI_JC1.GRID_COMBINAR Fg1, 0, 3, 0, 12, "DATOS DE LA OPERACIÓN", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC
        SGI_JC1.GRID_COMBINAR Fg1, 0, 13, 0, 16, "REFERENCIA", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC
        SGI_JC1.GRID_COMBINAR Fg1, 0, 17, 0, 19, "IMPORTE EN MN", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC
        SGI_JC1.GRID_COMBINAR Fg1, 0, 20, 0, 22, "IMPORTE EN ME", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC
        SGI_JC1.GRID_COMBINAR Fg1, 0, 23, 0, 24, "CUENTA", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC
        SGI_JC1.GRID_COMBINAR Fg1, 0, 25, 0, 26, "G/P", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC
      
        .TextMatrix(1, 1) = "IdDoc":                .ColWidth(1) = 0:  .ColAlignment(1) = flexAlignLeftCenter:   .Row = 1: .Col = 1: .CellAlignment = flexAlignLeftCenter
        
        .TextMatrix(1, 2) = "Sel":              .ColWidth(2) = 400:  .ColAlignment(2) = flexAlignLeftCenter:   .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 3) = "Num.Reg.":         .ColWidth(3) = 900:   .ColAlignment(3) = flexAlignLeftCenter:   .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 4) = "RUC":              .ColWidth(4) = 1100:  .ColAlignment(4) = flexAlignLeftCenter:   .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 5) = "Nombres":          .ColWidth(5) = 2000:  .ColAlignment(5) = flexAlignLeftCenter:   .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 6) = "T.D.":             .ColWidth(6) = 450:   .ColAlignment(6) = flexAlignLeftCenter:    .Row = 1: .Col = 6: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 7) = "Nº.Doc":           .ColWidth(7) = 1000:  .ColAlignment(7) = flexAlignLeftCenter:   .Row = 1: .Col = 7: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 8) = "Fch.Doc":          .ColWidth(8) = 800:   .ColAlignment(8) = flexAlignLeftCenter:   .Row = 1: .Col = 8: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 9) = "M":                .ColWidth(9) = 450:   .ColAlignment(9) = flexAlignLeftCenter:   .Row = 1: .Col = 9: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 10) = "T.C.":            .ColWidth(10) = 550:  .ColAlignment(10) = flexAlignRightCenter:  .Row = 1: .Col = 10: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 11) = "Importe":         .ColWidth(11) = 850:  .ColAlignment(11) = flexAlignRightCenter:  .Row = 1: .Col = 11: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 12) = "Glosa":           .ColWidth(12) = 400:  .ColAlignment(12) = flexAlignLeftCenter:  .Row = 1: .Col = 12: .CellAlignment = flexAlignLeftCenter
        
        .TextMatrix(1, 13) = "Mes":             .ColWidth(13) = 400:  .ColAlignment(13) = flexAlignLeftCenter:  .Row = 1: .Col = 13: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 14) = "Fch Ope":         .ColWidth(14) = 800:  .ColAlignment(14) = flexAlignCenterCenter: .Row = 1: .Col = 14: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 15) = "T.C.":            .ColWidth(15) = 550:  .ColAlignment(15) = flexAlignRightCenter:  .Row = 1: .Col = 15: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 16) = "Importe":         .ColWidth(16) = 850:  .ColAlignment(16) = flexAlignRightCenter:  .Row = 1: .Col = 16: .CellAlignment = flexAlignRightCenter
        
        .TextMatrix(1, 17) = "Debe MN":         .ColWidth(17) = 1000: .ColAlignment(17) = flexAlignRightCenter:   .Row = 1: .Col = 17: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 18) = "Haber MN":        .ColWidth(18) = 1000: .ColAlignment(18) = flexAlignRightCenter:   .Row = 1: .Col = 18: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 19) = "Saldo MN":        .ColWidth(19) = 1000: .ColAlignment(19) = flexAlignRightCenter:   .Row = 1: .Col = 19: .CellAlignment = flexAlignRightCenter
        
        .TextMatrix(1, 20) = "Debe ME":         .ColWidth(20) = 1000: .ColAlignment(20) = flexAlignRightCenter:   .Row = 1: .Col = 20: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 21) = "Haber ME":        .ColWidth(21) = 1000: .ColAlignment(21) = flexAlignRightCenter:   .Row = 1: .Col = 21: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 22) = "Saldo ME":        .ColWidth(22) = 1000: .ColAlignment(22) = flexAlignRightCenter:   .Row = 1: .Col = 22: .CellAlignment = flexAlignRightCenter
        
        .TextMatrix(1, 23) = "Num. Cta":        .ColWidth(23) = 900:  .ColAlignment(23) = flexAlignLeftCenter:    .Row = 1: .Col = 23: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 24) = "Nombre Cta":      .ColWidth(24) = 0:  .ColAlignment(24) = flexAlignLeftCenter:      .Row = 1: .Col = 24: .CellAlignment = flexAlignLeftCenter
         
        .TextMatrix(1, 25) = "Gan":             .ColWidth(25) = 400: .ColAlignment(25) = flexAlignCenterCenter:   .Row = 1: .Col = 25: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 26) = "Per":             .ColWidth(26) = 400:  .ColAlignment(26) = flexAlignCenterCenter:  .Row = 1: .Col = 26: .CellAlignment = flexAlignCenterCenter
              
        .TextMatrix(1, 27) = "IdCuen":
        .TextMatrix(1, 28) = "idprov":
        .TextMatrix(1, 29) = "Tipdoc":
        
        OCULTAR_COL Fg1, 27, 29
        
        Fg1.ColFormat(10) = "0.000"
        Fg1.ColFormat(15) = "0.000"
        
        Fg1.ColFormat(8) = SGI_JC1.FORMAT_DATE
        Fg1.ColFormat(14) = SGI_JC1.FORMAT_DATE
        
        'SGI_JC1.OCULTAR_COL fg1, 8, 12
        Fg1.ColDataType(2) = flexDTBoolean
        
    End With
    DoEvents
End Sub


'*******************************

Private Sub CmdBusProv_Click()
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
   
    Dim xCampos(5, 4) As String
    Dim nSQL As String
    
    xCampos(0, 0) = "Descripcion":      xCampos(0, 1) = "descripcion":  xCampos(0, 2) = "4500":        xCampos(0, 3) = "C"
    xCampos(1, 0) = "M":                xCampos(1, 1) = "simbolo":      xCampos(1, 2) = "450":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Cta Ganancia":     xCampos(2, 1) = "ctagnum":      xCampos(2, 2) = "1500":        xCampos(2, 3) = "C"
    xCampos(3, 0) = "Cta Perdida":      xCampos(3, 1) = "ctapnum":      xCampos(3, 2) = "1500":        xCampos(3, 3) = "C"
    xCampos(4, 0) = "Id":               xCampos(4, 1) = "id":           xCampos(4, 2) = "450":         xCampos(4, 3) = "N"
    
    Dim xObj As New SGI2_funciones.JC_Varios
    
    nSQL = "SELECT mae_ajuste.*, per.cuenta AS ctapnum, per.descripcion AS ctapdesc, gan.cuenta AS ctagnum, gan.descripcion AS ctagdesc, mae_moneda.simbolo " _
            + vbCr + " FROM ((mae_ajuste LEFT JOIN mae_moneda ON mae_ajuste.idmon = mae_moneda.id) LEFT JOIN con_planctas AS per ON mae_ajuste.idcuenper = per.id) LEFT JOIN con_planctas AS gan ON mae_ajuste.idcuengan = gan.id"
    
    xObj.CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Redondeo a Centimos", "descripcion", "descripcion", Principio
    Set xObj = Nothing
    If xRs.State = 1 Then
        TxtAjuste.Text = NulosC(xRs("descripcion"))
        LblIdAjuste.Caption = NulosC(xRs("idmon"))
        
        LblIdLibro.Caption = NulosN(xRs("idlib"))
        
        LblIdMon.Caption = NulosC(xRs("idmon"))
        
        '--cuenta ganancia
        TxtGanancia.Text = NulosC(xRs("ctagnum"))
        LblDescGanancia.Caption = NulosC(xRs("ctagdesc"))
        LblIdCtaGan.Caption = NulosN(xRs("idcuengan"))
        '--cuenta perdida
        TxtPerdida.Text = NulosC(xRs("ctapnum"))
        LblDescPerdida.Caption = NulosC(xRs("ctapdesc"))
        LblIdCtaPer.Caption = NulosN(xRs("idcuenper"))

    End If
    
    Set xRs = Nothing
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If OptSel(0).Value = True Then
        If Button.Index = 1 Then pConsultar
        If Button.Index = 2 Then Grabar
    Else
        '--bancos
        If Button.Index = 1 Then pConsultar1
        If Button.Index = 2 Then Grabar1
    End If
    
    If Button.Index = 4 Then pExportar
'    If Button.Index = 4 Then pImprimir
    If Button.Index = 6 Then
        Unload Me
    End If
End Sub

Private Sub TxtLibro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtLibro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusProv_Click
    End If
End Sub




Private Sub pExportar()
'    Dim xFun As New SGI2_funciones.formularios
'    Dim rst As New ADODB.Recordset
    
'    If fg1.Rows = fg1.FixedRows Then
'        MsgBox "No hay registro para exportar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Exit Sub
'    End If
'    xFun.VSFlexGrid_Exportar_MSExcel xCon, fg1, "CONSULTA DE ASIENTO Nº. " & NumRegistro, "", "", "Consulta de Asiento"
'
    GRID_EXPORTAR_MSEXCELTMP Fg1, xCon, flexFileCustomText, True, "ORIGEN PARA DIFERENCIA DE CAMBIO", "DEL 01/01/09 AL 31/12/09", " "

    
'    Set xFun = Nothing
    
End Sub


Private Sub pImprimir()

    Dim xPrint As New SGI2_funciones.formularios
    
    Me.MousePointer = vbHourglass
    xPrint.Imprimir_x_VSFlexGrid Fg1, "CONSULTA DE ASIENTO Nº. " & NumRegistro, " ", "", False, True
    Set xPrint = Nothing
    Me.MousePointer = vbDefault
  
    
End Sub

Function Grabar() As Boolean

'--Modificado 23/09/11 Johan Castro
'            Asignar valores de cuenta a variables, Validar el ingreso de cuentas que sean diferentes a cero

    Dim A, B, Rpta As Integer
    Dim RstDia As New ADODB.Recordset
    
    Dim xIdCuen, xId As Integer
    Dim xTotal As Double
    Dim xNumAsiento As String
    Dim xSaldo As Double '--indica el saldo actual del documento
    Dim mRow&
    Dim mIdMes As Integer
    
    Dim IdCtaGanancia As Long
    Dim IdCtaPerdida As Long
    Dim TipCam As Double
    
    
    Dim IdCtaDestDeb As Long
    Dim IdCtaDestHab As Long
    
    Dim sSaldo As Double
    Dim nSQL As String
    Dim mIdMov As Double
    
    Dim mAjusteMoneda As Integer '--indica la moneda de ajuste por diferencia de cambio
                                 '=1; Cuando Ajuste es de ME a MN
                                 '=0; Cuando Ajuste es de MN a ME
    
    Dim rst As New ADODB.Recordset
    
    On Error GoTo LaCague
    
    
    '--asignando valores a las variables
    IdCtaGanancia = NulosN(LblIdCtaGan.Caption)
    IdCtaPerdida = NulosN(LblIdCtaPer.Caption)
    '--validando la seleccion de las cuentas x diferencia de cambio
    If IdCtaGanancia = 0 Or IdCtaPerdida = 0 Then
        MsgBox "Seleccione la cuenta de Ganáncia x Diferencia de Cambio", vbInformation
        TxtGanancia.SetFocus
        Exit Function
    End If
    If IdCtaPerdida = 0 Then
        MsgBox "Seleccione la cuenta de Pérdida x Diferencia de Cambio", vbInformation
        TxtPerdida.SetFocus
        Exit Function
    End If
    '------------
    
    
    nSQL = "SELECT con_planctas.ctadesdeb, con_planctas.ctadeshab FROM con_planctas WHERE (((con_planctas.id)=" & IdCtaPerdida & "));"
    RST_Busq rst, nSQL, xCon
    If rst.RecordCount <> 0 Then
        IdCtaDestDeb = NulosN(rst("ctadesdeb"))
        IdCtaDestHab = NulosN(rst("ctadeshab"))
    Else
        IdCtaDestDeb = 0
        IdCtaDestHab = 0
    End If
    Set rst = Nothing
    
    '--obteniendo el ultimo id de movimiento de ajuste por diferencia de cambio
    nSQL = "SELECT con_diario.idlib, Last(con_diario.idmov) AS idmov1 FROM con_diario GROUP BY con_diario.idlib HAVING (((con_diario.idlib)=44));"
    RST_Busq rst, nSQL, xCon
    If rst.RecordCount <> 0 Then
        mIdMov = NulosN(rst("idmov1"))
    Else
        mIdMov = 0
    End If
    Set rst = Nothing
    
    
    RST_Busq RstDia, "select top 1 * from con_diario", xCon
    
    xCon.BeginTrans
    
    '--eliminar los registros del libro seleccionado
    nSQL = "delete from con_diario where idlib=44 and idmov in ( select idmov from con_diario where idlib=44 and ridlib=" & NulosN(LblIdLibro.Caption) & " and idmon= " & NulosN(LblIdMon.Caption) & ")"
    xCon.Execute nSQL
    'xCon.Execute "delete from con_diario where idlib = 44 "
    
    '--identificar el ajuste en que moneda se va realizar
    If NulosN(LblIdMon.Caption) = 1 Then
        mAjusteMoneda = 2
    Else
        mAjusteMoneda = 1
    End If
    '----------------------------------------------
    
    Frame1.Visible = True
    ProgressBar2.Max = Fg1.Rows - 1
    ProgressBar2.Min = 0
    
    For mRow = 2 To Fg1.Rows - 1
    
        ProgressBar2.Value = mRow
        DoEvents
        
        mIdMov = mIdMov + 1
        
        mIdMes = Month(Fg1.TextMatrix(mRow, 8))
        sSaldo = NulosN(Fg1.TextMatrix(mRow, 19))
        xNumAsiento = NuevoNumAsiento(44, mIdMes, xCon)
        TipCam = NulosN(Fg1.TextMatrix(mRow, 15))
        '**************************************************************************************
        
        If sSaldo = 0 Then
            MsgBox "SSS"
        End If
        
        '--agregando la ganancia o perdida en diario
        If LCase(Fg1.TextMatrix(mRow, 25)) <> "" Or LCase(Fg1.TextMatrix(mRow, 26)) <> "" Then
            RstDia.AddNew
            RstDia("año") = AnoTra
            RstDia("idmes") = mIdMes
            RstDia("idlib") = 44
            RstDia("idmov") = mIdMov
            RstDia("numasi") = xNumAsiento
            RstDia("tc") = TipCam
            RstDia("iddoc") = NulosN(Fg1.TextMatrix(mRow, 1))
            
            If LCase(Fg1.TextMatrix(mRow, 25)) = "si" Then
                RstDia("idcue") = IdCtaGanancia
                RstDia("impdebsol") = 0
                RstDia("impdebdol") = 0
                RstDia("imphabsol") = Abs(sSaldo)
                RstDia("imphabdol") = Abs(sSaldo / TipCam)
                '--poner nombre cuenta en glosa
                RstDia("rglosaope") = Busca_Codigo(IdCtaGanancia, "id", "descripcion", "con_planctas", "N", xCon)     'ok
                
            Else
                RstDia("idcue") = IdCtaPerdida
                RstDia("impdebsol") = Abs(sSaldo)
                RstDia("impdebdol") = Abs(sSaldo / TipCam)
                RstDia("imphabsol") = 0
                RstDia("imphabdol") = 0
                '--poner nombre cuenta en glosa
                RstDia("rglosaope") = Busca_Codigo(IdCtaPerdida, "id", "descripcion", "con_planctas", "N", xCon)     'ok
            End If
            
            RstDia("fchasi") = CDate("01/" + Format(mIdMes, "00") + "/" + AnoTra)
            RstDia("fchdoc") = CDate(Fg1.TextMatrix(mRow, 14))
            
            RstDia("idmon") = NulosN(LblIdMon.Caption)
            Select Case NulosN(LblIdLibro.Caption)
                Case 2, 37, 41, 42
                    RstDia("ridlib") = NulosN(LblIdLibro.Caption)
                Case Else
                    RstDia("ridlib") = 0
            End Select
            RstDia("iddocpro") = 0
            RstDia("ridtipper") = mIdTipPer
            RstDia("ridper") = NulosN(Fg1.TextMatrix(mRow, 28))
            RstDia("rtipdoc") = NulosN(Fg1.TextMatrix(mRow, 29))
            RstDia("rfchope") = CDate(Fg1.TextMatrix(mRow, 8))
            RstDia("rnumerodoc") = Fg1.TextMatrix(mRow, 7)
            RstDia("rregistro") = Fg1.TextMatrix(mRow, 3)
            RstDia("rglosa") = Fg1.TextMatrix(mRow, 12)
            
            
            RstDia("rglosa") = Fg1.TextMatrix(mRow, 12)
            RstDia("ridmon") = NulosN(LblIdMon.Caption)
            RstDia("ajuste") = mAjusteMoneda
            RstDia("aplicatc") = -1 '--considerar tc de diario
            RstDia.Update
                        
            '**************************************************************************************
            '--provicion
            RstDia.AddNew
            RstDia("año") = AnoTra
            RstDia("idmes") = mIdMes
            RstDia("idlib") = 44
            RstDia("idmov") = mIdMov
            RstDia("numasi") = xNumAsiento
            RstDia("tc") = TipCam
            RstDia("iddoc") = NulosN(Fg1.TextMatrix(mRow, 1))
            
            RstDia("idcue") = NulosN(Fg1.TextMatrix(mRow, 27))
            
            RstDia("fchasi") = CDate("01/" + Format(mIdMes, "00") + "/" + AnoTra)
            RstDia("fchdoc") = CDate(Fg1.TextMatrix(mRow, 14))
    
                
            If LCase(Fg1.TextMatrix(mRow, 25)) = "si" Then
                RstDia("impdebsol") = Abs(sSaldo)
                RstDia("impdebdol") = Abs(sSaldo / TipCam)
    
                RstDia("imphabsol") = 0
                RstDia("imphabdol") = 0
                '--poner nombre cuenta en glosa
                RstDia("rglosaope") = Busca_Codigo(IdCtaGanancia, "id", "descripcion", "con_planctas", "N", xCon)     'ok
            Else
                RstDia("imphabsol") = 0
                RstDia("imphabdol") = 0
                
                RstDia("imphabsol") = Abs(sSaldo)
                RstDia("imphabdol") = Abs(sSaldo / TipCam)
                '--poner nombre cuenta en glosa
                RstDia("rglosaope") = Busca_Codigo(IdCtaPerdida, "id", "descripcion", "con_planctas", "N", xCon)     'ok
            End If
            
            RstDia("idmon") = NulosN(LblIdMon.Caption)
            Select Case NulosN(LblIdLibro.Caption)
                Case 1, 40, 9, 999
                    RstDia("ridlib") = NulosN(LblIdLibro.Caption)
                Case Else
                    RstDia("ridlib") = 0
            End Select
            RstDia("iddocpro") = 0
            RstDia("ridtipper") = mIdTipPer
            RstDia("ridper") = Fg1.TextMatrix(mRow, 28)
            RstDia("rfchope") = CDate(Fg1.TextMatrix(mRow, 14))
            
            RstDia("rtipdoc") = NulosN(Fg1.TextMatrix(mRow, 29))
            RstDia("rnumerodoc") = Fg1.TextMatrix(mRow, 7)
            RstDia("rregistro") = Fg1.TextMatrix(mRow, 3)
            RstDia("rglosa") = Fg1.TextMatrix(mRow, 12)
            RstDia("ridmon") = NulosN(LblIdMon.Caption)
            
            RstDia("ajuste") = mAjusteMoneda
            RstDia("aplicatc") = -1 '--considerar tc de diario
            RstDia.Update
            
            '**************************************************************************************
        
        
            '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            '--destinos de perdida
           
            If IdCtaDestDeb <> 0 And IdCtaDestHab <> 0 And LCase(Fg1.TextMatrix(mRow, 26)) = "si" Then
                '************************************************************************************************
                '--destinos automatico cta debe
                RstDia.AddNew
                RstDia("año") = AnoTra
                RstDia("idmes") = mIdMes
                RstDia("idlib") = 44
                RstDia("idmov") = mIdMov
                RstDia("numasi") = xNumAsiento
                RstDia("tc") = TipCam
                RstDia("iddoc") = NulosN(Fg1.TextMatrix(mRow, 1))
                RstDia("impdebsol") = 0
                RstDia("impdebdol") = 0
                RstDia("imphabsol") = 0
                RstDia("imphabdol") = 0
                RstDia("idcue") = IdCtaDestDeb
                If NulosN(LblIdMon.Caption) = 2 Then
                    RstDia("impdebdol") = Abs(sSaldo) / TipCam
                Else
                    RstDia("impdebsol") = Abs(sSaldo) * TipCam
                End If
                
                RstDia("fchasi") = CDate("01/" + Format(mIdMes, "00") + "/" + AnoTra)
                RstDia("fchdoc") = CDate(Fg1.TextMatrix(mRow, 14))
                RstDia("idmon") = NulosN(LblIdMon.Caption)

                RstDia("ridlib") = 0
                RstDia("iddocpro") = 0
                RstDia("ridtipper") = mIdTipPer
                RstDia("ridper") = Fg1.TextMatrix(mRow, 28)
                RstDia("rfchope") = CDate(Fg1.TextMatrix(mRow, 14))
                
                RstDia("rtipdoc") = NulosN(Fg1.TextMatrix(mRow, 29))
                RstDia("rnumerodoc") = Fg1.TextMatrix(mRow, 7)
                RstDia("rregistro") = Fg1.TextMatrix(mRow, 3)
                RstDia("rglosa") = Fg1.TextMatrix(mRow, 12)
                RstDia("ridmon") = NulosN(LblIdMon.Caption)
                
                RstDia("ajuste") = mAjusteMoneda
                RstDia("aplicatc") = -1 '--considerar tc de diario
                
                RstDia.Update
                
                '************************************************************************************************
                '--destinos automatico cta debe
                RstDia.AddNew
                RstDia("año") = AnoTra
                RstDia("idmes") = mIdMes
                RstDia("idlib") = 44
                RstDia("idmov") = mIdMov
                RstDia("numasi") = xNumAsiento
                RstDia("tc") = TipCam
                RstDia("iddoc") = NulosN(Fg1.TextMatrix(mRow, 1))
                RstDia("impdebsol") = 0
                RstDia("impdebdol") = 0
                RstDia("imphabsol") = 0
                RstDia("imphabdol") = 0
                RstDia("idcue") = IdCtaDestHab
                
                If NulosN(LblIdMon.Caption) = 2 Then
                    RstDia("imphabdol") = Abs(sSaldo) / TipCam
                Else
                    RstDia("imphabsol") = Abs(sSaldo) * TipCam
                End If
                
                RstDia("fchasi") = CDate("01/" + Format(mIdMes, "00") + "/" + AnoTra)
                RstDia("fchdoc") = CDate(Fg1.TextMatrix(mRow, 14))
                RstDia("idmon") = NulosN(LblIdMon.Caption)

                RstDia("ridlib") = 0
                RstDia("iddocpro") = 0
                RstDia("ridtipper") = mIdTipPer
                RstDia("ridper") = Fg1.TextMatrix(mRow, 28)
                RstDia("rfchope") = CDate(Fg1.TextMatrix(mRow, 14))
                
                RstDia("rtipdoc") = NulosN(Fg1.TextMatrix(mRow, 29))
                RstDia("rnumerodoc") = Fg1.TextMatrix(mRow, 7)
                RstDia("rregistro") = Fg1.TextMatrix(mRow, 3)
                RstDia("rglosa") = Fg1.TextMatrix(mRow, 12)
                RstDia("ridmon") = NulosN(LblIdMon.Caption)
                
                RstDia("ajuste") = mAjusteMoneda
                RstDia("aplicatc") = -1 '--considerar tc de diario

                RstDia.Update
                '************************************************************************************************
                
            End If
            '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        
        End If
        
        
    Next

    xCon.CommitTrans
    
    MsgBox "El proceso termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo

    Set RstDia = Nothing
    
    Grabar = True
    
    Frame1.Visible = False
    
    Exit Function
    
LaCague:
'    Resume
    xCon.RollbackTrans
    Set RstDia = Nothing
    Frame1.Visible = False
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" & vbCr & Trim(Err.Description)
End Function



'--AJUSTE BANCOS


Private Sub Configurar_Grilla1()

    With Fg1
        '-----
        .Rows = 2
        .FixedRows = 2
        .Cols = 15
        
        .ColWidth(0) = 200
        '--DATOS DE FILA
        
        SGI_JC1.GRID_COMBINAR Fg1, 0, 3, 0, 8, "DATOS DE LA OPERACIÓN", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC
        SGI_JC1.GRID_COMBINAR Fg1, 0, 9, 0, 11, "IMPORTE", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC
        SGI_JC1.GRID_COMBINAR Fg1, 0, 12, 0, 13, "G/P", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC
      
        .TextMatrix(1, 1) = "IdDoc":        .ColWidth(1) = 0:     .ColAlignment(1) = flexAlignLeftCenter:   .Row = 1: .Col = 1: .CellAlignment = flexAlignLeftCenter
        
        .TextMatrix(1, 2) = "Sel":          .ColWidth(2) = 400:   .ColAlignment(2) = flexAlignLeftCenter:   .Row = 1: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 3) = "Num.Reg.":     .ColWidth(3) = 900:   .ColAlignment(3) = flexAlignLeftCenter:   .Row = 1: .Col = 3: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 4) = "Tipo":         .ColWidth(4) = 800:   .ColAlignment(4) = flexAlignLeftCenter:   .Row = 1: .Col = 4: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 5) = "Fch.Doc":      .ColWidth(5) = 800:   .ColAlignment(5) = flexAlignLeftCenter:   .Row = 1: .Col = 5: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 6) = "M":            .ColWidth(6) = 450:   .ColAlignment(6) = flexAlignLeftCenter:   .Row = 1: .Col = 6: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 7) = "T.C.":         .ColWidth(7) = 550:   .ColAlignment(7) = flexAlignRightCenter:  .Row = 1: .Col = 7: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 8) = "Glosa":        .ColWidth(8) = 2500:   .ColAlignment(8) = flexAlignLeftCenter:   .Row = 1: .Col = 8: .CellAlignment = flexAlignLeftCenter
        
        .TextMatrix(1, 9) = "Debe":         .ColWidth(9) = 1000:  .ColAlignment(9) = flexAlignRightCenter:    .Row = 1: .Col = 9: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 10) = "Haber":       .ColWidth(10) = 1000: .ColAlignment(10) = flexAlignRightCenter:   .Row = 1: .Col = 10: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 11) = "Saldo":       .ColWidth(11) = 1000: .ColAlignment(11) = flexAlignRightCenter:   .Row = 1: .Col = 11: .CellAlignment = flexAlignRightCenter
        
        .TextMatrix(1, 12) = "Gan":         .ColWidth(12) = 400:  .ColAlignment(12) = flexAlignCenterCenter:  .Row = 1: .Col = 12: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 13) = "Per":         .ColWidth(13) = 400:  .ColAlignment(13) = flexAlignCenterCenter:  .Row = 1: .Col = 13: .CellAlignment = flexAlignCenterCenter
              
        .TextMatrix(1, 14) = "idmes":
        
        OCULTAR_COL Fg1, 14, 14
        
        Fg1.ColFormat(7) = "0.000"
        
        Fg1.ColFormat(9) = SGI_JC1.FORMAT_MONTO
        Fg1.ColFormat(10) = SGI_JC1.FORMAT_MONTO
        Fg1.ColFormat(11) = SGI_JC1.FORMAT_MONTO
        
        Fg1.ColFormat(5) = SGI_JC1.FORMAT_DATE
        
        Fg1.ColDataType(2) = flexDTBoolean
        
    End With
    DoEvents
End Sub




Private Sub pConsultar1()
    Dim nSQL As String
    Dim rst As New ADODB.Recordset
    
    Configurar_Grilla1
    
    DoEvents
    
    If NulosN(LblIdAjuste.Caption) = 0 Then
        MsgBox "Seleccione el Tipo de Diferencia de Cambio", vbExclamation, xTitulo
        TxtAjuste.SetFocus
        Exit Sub
    End If
    
    '--cargando la lista de pagos
    
    'En ME
    If NulosN(LblIdMon.Caption) = 2 Then
            
        nSQL = "SELECT banco.idlib, banco.idmes, banco.idmov, banco.idmon, banco.registro,banco.tipo1, banco.fchope, banco.glosa, Last(banco.tipcam) AS tipcam, banco.simbolo, Sum(banco.impdebesol) AS totdeb, Sum(banco.imphabersol) AS tothab, Sum(banco.impdebedol) AS totdebdol, Sum(banco.imphaberdol) AS tothabdol, Sum(banco.impdebesol)-Sum(banco.imphabersol) AS totsal, Sum(banco.impdebedol)-Sum(banco.imphaberdol) AS saldol " _
                + vbCr + " FROM ( " _
                + vbCr + " SELECT con_diario.idlib, con_diario.idmes, con_diario.idmov, con_diario.idmon, Format(con_diario.idmes,'00') & IIf(mae_libros.codsun Is Null,'',mae_libros.codsun) & con_diario.numasi AS registro,mae_tipomov.descripcion AS tipo1, mae_libros.descripcion AS libdesc, con_diario.fchdoc AS fchope, con_diario.rglosaope as glosa, iif(con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven)) AS tipcam, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo, " _
                + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.impdebdol*tipcam),con_diario.impdebsol) AS impdebesol, " _
                + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.imphabdol*tipcam),con_diario.imphabsol) AS imphabersol, " _
                + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(tipcam=0 Or con_diario.impdebsol=0,0,(con_diario.impdebsol/tipcam))) AS impdebedol, " _
                + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(tipcam=0 Or con_diario.imphabsol=0,0,(con_diario.imphabsol/tipcam))) As imphaberdol " _
                + vbCr + " FROM ((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) INNER JOIN (tes_caja INNER JOIN mae_tipomov ON tes_caja.tipmov = mae_tipomov.id) ON con_diario.idmov = tes_caja.id " _
                + vbCr + " WHERE (((con_diario.idlib)=6) AND ((con_diario.ajuste) In (0,2)) AND ((con_diario.fchasi) Between CDate('01/01/" & AnoTra & "') And CDate('31/12/" & AnoTra & "'))) " _
                + vbCr + " ) AS banco " _
                + vbCr + " GROUP BY banco.idlib, banco.idmes, banco.idmov, banco.idmon, banco.registro,banco.tipo1, banco.fchope, banco.glosa, banco.simbolo " _
                + vbCr + " HAVING (((banco.idmon)=2) AND ((Sum([banco].[impdebesol])-Sum([banco].[imphabersol])) Not Between -0.001 And 0.001) AND ((Sum([banco].[impdebedol])-Sum([banco].[imphaberdol])) Between -0.001 And 0.001)) "

    Else
        '--En MN
    
        nSQL = "SELECT banco.idlib, banco.idmes, banco.idmov, banco.idmon, banco.registro,banco.tipo1, banco.fchope, banco.glosa, Last(banco.tipcam) AS tipcam, banco.simbolo, Sum(banco.impdebesol) AS totdebsol, Sum(banco.imphabersol) AS tothabsol, Sum(banco.impdebedol) AS totdeb, Sum(banco.imphaberdol) AS tothab, Sum(banco.impdebesol)-Sum(banco.imphabersol) AS salsol, Sum(banco.impdebedol)-Sum(banco.imphaberdol) AS totsal " _
                + vbCr + " FROM ( " _
                + vbCr + " SELECT con_diario.idlib, con_diario.idmes, con_diario.idmov, con_diario.idmon, Format(con_diario.idmes,'00') & IIf(mae_libros.codsun Is Null,'',mae_libros.codsun) & con_diario.numasi AS registro, mae_tipomov.descripcion AS tipo1,mae_libros.descripcion AS libdesc, con_diario.fchdoc AS fchope, con_diario.rglosaope as glosa, iif(con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven)) AS tipcam, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo, " _
                + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.impdebdol*tipcam),con_diario.impdebsol) AS impdebesol, " _
                + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.imphabdol*tipcam),con_diario.imphabsol) AS imphabersol, " _
                + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(tipcam=0 Or con_diario.impdebsol=0,0,(con_diario.impdebsol/tipcam))) AS impdebedol, " _
                + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(tipcam=0 Or con_diario.imphabsol=0,0,(con_diario.imphabsol/tipcam))) As imphaberdol " _
                + vbCr + " FROM ((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) INNER JOIN (tes_caja INNER JOIN mae_tipomov ON tes_caja.tipmov = mae_tipomov.id) ON con_diario.idmov = tes_caja.id " _
                + vbCr + " WHERE (((con_diario.idlib)=6) AND ((con_diario.ajuste) In (0,1)) AND ((con_diario.fchasi) Between CDate('01/01/" & AnoTra & "') And CDate('31/12/" & AnoTra & "'))) " _
                + vbCr + " ) AS banco " _
                + vbCr + " GROUP BY banco.idlib, banco.idmes, banco.idmov, banco.idmon, banco.registro,banco.tipo1, banco.fchope, banco.glosa, banco.simbolo " _
                + vbCr + " HAVING (((banco.idmon)=1) AND ((Sum([banco].[impdebedol])-Sum([banco].[imphaberdol])) Not Between -0.001 And 0.001) AND ((Sum([banco].[impdebesol])-Sum([banco].[imphabersol])) Between -0.001 And 0.001)) "
    
    End If
    Me.MousePointer = vbHourglass
    RST_Busq rst, nSQL, xCon
    
    Fg1.Rows = Fg1.FixedRows
    DoEvents
    
    Frame1.Visible = True
    ProgressBar2.Min = 0
    ProgressBar2.Value = 0
    
    If rst.RecordCount <> 0 Then
    ProgressBar2.Max = rst.RecordCount
    
        Do While Not rst.EOF
            DoEvents
            ProgressBar2.Value = ProgressBar2.Value + 1
            Fg1.Rows = Fg1.Rows + 1
        
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = rst("idmov")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = -1
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(rst("registro"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(rst("tipo1"))
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(rst("fchope"))
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(rst("simbolo"))
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosC(rst("tipcam"))
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosC(rst("glosa"))
            
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(NulosC(rst("totdeb")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(NulosC(rst("tothab")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(NulosN(rst("totsal")), FORMAT_MONTO)
            
            
            Fg1.TextMatrix(Fg1.Rows - 1, 14) = NulosN(rst("idmes"))
            
            If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 11)) > 0 Then '--
                Fg1.TextMatrix(Fg1.Rows - 1, 12) = "Si"
                Fg1.TextMatrix(Fg1.Rows - 1, 13) = ""
            Else
                Fg1.TextMatrix(Fg1.Rows - 1, 12) = ""
                Fg1.TextMatrix(Fg1.Rows - 1, 13) = "Si"
            End If
            
Seguir:
            
            rst.MoveNext
        Loop
        
    
    End If
    Frame1.Visible = False
    Set rst = Nothing
    
    Me.MousePointer = vbDefault

End Sub




Function Grabar1() As Boolean
    
    Dim RstDia As New ADODB.Recordset
    
    Dim xIdCuen, xId As Integer
    Dim xTotal As Double
    Dim xNumAsiento As String
    Dim xSaldo As Double '--indica el saldo actual del documento
    Dim mRow&
    Dim mIdMes As Integer
    Dim mIdMov As Double
    Dim IdCtaGanancia As Long
    Dim IdCtaPerdida As Long
    
    Dim IdCtaDestDeb As Long
    Dim IdCtaDestHab As Long
    
    Dim nSQL As String
    
    Dim TipCam As Double
    
    
    Dim sSaldo As Double
    
    
    Dim rst As New ADODB.Recordset
    
    On Error GoTo LaCague
    
    RST_Busq rst, "SELECT mae_ajuste.* FROM mae_ajuste WHERE mae_ajuste.idmon=2 and mae_ajuste.idlib = 6 ;", xCon
    IdCtaGanancia = NulosN(rst("idcuengan"))
    IdCtaPerdida = NulosN(rst("idcuenper"))
    
    Set rst = Nothing
    
    nSQL = "SELECT con_planctas.ctadesdeb, con_planctas.ctadeshab FROM con_planctas WHERE (((con_planctas.id)=" & IdCtaPerdida & "));"
    RST_Busq rst, nSQL, xCon
    If rst.RecordCount <> 0 Then
        IdCtaDestDeb = NulosN(rst("ctadesdeb"))
        IdCtaDestHab = NulosN(rst("ctadeshab"))
    Else
        IdCtaDestDeb = 0
        IdCtaDestHab = 0
    End If
    Set rst = Nothing
    
    RST_Busq RstDia, "select top 1 * from con_diario", xCon
    
    xCon.BeginTrans
    
'    If NulosN(LblIdMon.Caption) = 2 Then
        xCon.Execute "delete from con_diario where ajuste = 1 "
'    Else
'        xCon.Execute "delete from con_diario where ajuste = 2 "
'    End If
    
    
    Frame1.Visible = True
    ProgressBar2.Max = Fg1.Rows - 1
    ProgressBar2.Min = 0
    
    For mRow = 2 To Fg1.Rows - 1
        ProgressBar2.Value = mRow
        DoEvents
        
        mIdMes = NulosN(Fg1.TextMatrix(mRow, 14))
        sSaldo = NulosN(Fg1.TextMatrix(mRow, 11))
        mIdMov = NulosN(Fg1.TextMatrix(mRow, 1))
        
        xNumAsiento = DevuelveNumAsiento(6, mIdMov, mIdMes, xCon)
        TipCam = NulosN(Fg1.TextMatrix(mRow, 7))
        
        '**************************************************************************************
        If sSaldo = 0 Then
            MsgBox "SSS"
        End If
        
        RstDia.AddNew
        RstDia("año") = AnoTra
        RstDia("idmes") = mIdMes
        RstDia("idlib") = 6
        RstDia("idmov") = mIdMov
        RstDia("numasi") = xNumAsiento
        RstDia("tc") = TipCam
        
        RstDia("impdebsol") = 0
        RstDia("impdebdol") = 0
        RstDia("imphabsol") = 0
        RstDia("imphabdol") = 0
                
        If NulosN(LblIdMon.Caption) = 2 Then
            If LCase(Fg1.TextMatrix(mRow, 12)) = "si" Then
                RstDia("idcue") = IdCtaGanancia
                RstDia("imphabdol") = Abs(sSaldo) / TipCam
                
            Else
                RstDia("idcue") = IdCtaPerdida
                RstDia("impdebdol") = Abs(sSaldo) / TipCam
            End If
        
        Else
            If LCase(Fg1.TextMatrix(mRow, 12)) = "si" Then
                RstDia("idcue") = IdCtaGanancia
                RstDia("imphabsol") = Abs(sSaldo) * TipCam
            Else
                RstDia("idcue") = IdCtaPerdida
                RstDia("impdebsol") = Abs(sSaldo) * TipCam
            End If
        End If
        
        RstDia("fchasi") = CDate("01/" + Format(mIdMes, "00") + "/" + AnoTra)
        RstDia("fchdoc") = CDate(Fg1.TextMatrix(mRow, 5))
        
        RstDia("idmon") = NulosN(LblIdMon.Caption)
        
        RstDia("ridlib") = 6
        RstDia("iddocpro") = 0
        RstDia("ridtipper") = 9999
        RstDia("ridper") = 0
        RstDia("rtipdoc") = 0
        RstDia("rfchope") = Null
        RstDia("rnumerodoc") = Null
        RstDia("rregistro") = Fg1.TextMatrix(mRow, 3)
        RstDia("rglosaope") = Fg1.TextMatrix(mRow, 8)
        RstDia("ridmon") = NulosN(LblIdMon.Caption)
        RstDia("ajuste") = 1
        RstDia("aplicatc") = -1
        RstDia.Update
        '**************************************************************************************
        '--destinos de perdida
        
        If IdCtaDestDeb <> 0 And IdCtaDestHab <> 0 And LCase(Fg1.TextMatrix(mRow, 13)) = "si" Then
            '************************************************************************************************
            '--destinos automatico cta debe
            RstDia.AddNew
            RstDia("año") = AnoTra
            RstDia("idmes") = mIdMes
            RstDia("idlib") = 6
            RstDia("idmov") = mIdMov
            RstDia("numasi") = xNumAsiento
            RstDia("tc") = TipCam
            
            RstDia("impdebsol") = 0
            RstDia("impdebdol") = 0
            RstDia("imphabsol") = 0
            RstDia("imphabdol") = 0
            RstDia("idcue") = IdCtaDestDeb
            If NulosN(LblIdMon.Caption) = 2 Then
                RstDia("impdebdol") = Abs(sSaldo) / TipCam
            Else
                RstDia("impdebsol") = Abs(sSaldo) * TipCam
            End If
            RstDia("fchasi") = CDate("01/" + Format(mIdMes, "00") + "/" + AnoTra)
            RstDia("fchdoc") = CDate(Fg1.TextMatrix(mRow, 5))
            RstDia("idmon") = NulosN(LblIdMon.Caption)
            RstDia("ridlib") = 6
            RstDia("iddocpro") = 0
            RstDia("ridtipper") = 9999
            RstDia("ridper") = 0
            RstDia("rtipdoc") = 0
            RstDia("rfchope") = Null
            RstDia("rnumerodoc") = Null
            RstDia("rregistro") = Fg1.TextMatrix(mRow, 3)
            RstDia("rglosaope") = Fg1.TextMatrix(mRow, 8)
            RstDia("ridmon") = NulosN(LblIdMon.Caption)
            RstDia("ajuste") = 1
            RstDia("aplicatc") = -1
            RstDia.Update
            
            '************************************************************************************************
            '--destinos automatico cta debe
            RstDia.AddNew
            RstDia("año") = AnoTra
            RstDia("idmes") = mIdMes
            RstDia("idlib") = 6
            RstDia("idmov") = mIdMov
            RstDia("numasi") = xNumAsiento
            RstDia("tc") = TipCam
            
            RstDia("impdebsol") = 0
            RstDia("impdebdol") = 0
            RstDia("imphabsol") = 0
            RstDia("imphabdol") = 0
            RstDia("idcue") = IdCtaDestHab
            
            If NulosN(LblIdMon.Caption) = 2 Then
                RstDia("imphabdol") = Abs(sSaldo) / TipCam
            Else
                RstDia("imphabsol") = Abs(sSaldo) * TipCam
            End If
            RstDia("fchasi") = CDate("01/" + Format(mIdMes, "00") + "/" + AnoTra)
            RstDia("fchdoc") = CDate(Fg1.TextMatrix(mRow, 5))
            RstDia("idmon") = NulosN(LblIdMon.Caption)
            RstDia("ridlib") = 6
            RstDia("iddocpro") = 0
            RstDia("ridtipper") = 9999
            RstDia("ridper") = 0
            RstDia("rtipdoc") = 0
            RstDia("rfchope") = Null
            RstDia("rnumerodoc") = Null
            RstDia("rregistro") = Fg1.TextMatrix(mRow, 3)
            RstDia("rglosaope") = Fg1.TextMatrix(mRow, 8)
            RstDia("ridmon") = NulosN(LblIdMon.Caption)
            RstDia("ajuste") = 1
            RstDia("aplicatc") = -1
            RstDia.Update
            '************************************************************************************************
            
        End If
    Next

    xCon.CommitTrans
    
    MsgBox "El proceso terminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo

    Set RstDia = Nothing
    
    
    
    Frame1.Visible = False
    
    Exit Function
    
LaCague:
'    Resume
    xCon.RollbackTrans
    Set RstDia = Nothing
    Frame1.Visible = False
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" & vbCr & Trim(Err.Description)
End Function






Private Sub pConsultarantes()

    Dim nSQL As String
    Dim rst As New ADODB.Recordset
    Dim rstPago As New ADODB.Recordset
    
    '--cargando datos de provicion
    nSQL = "SELECT com_compras.id, com_compras.idmon,com_compras.idpro,com_compras.tipdoc,  com_compras.glosa,mae_prov.numruc, [mae_prov]![nombre] AS nombre, IIf([com_compras].[numreg] Is Null Or [com_compras].[numreg]='','',Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Right([com_compras].[numreg],4)) AS registro, 'Compras' AS libro, mae_documento.abrev, IIf([com_compras]![numser] Is Null Or [com_compras]![numser]='','',[com_compras]![numser]+'-')+[com_compras]![numdoc] AS numdoc2, " _
        + vbCr + " com_compras.fchdoc, mae_moneda.simbolo, con_tc.impven AS tipcam, com_compras.imptot AS imptotal, IIf([com_compras].[imptot] Is Null,0,IIf([com_compras].[idmon]=1,[com_compras].[imptot],IIf([con_tc].[impven] Is Null,0,[com_compras].[imptot]*[con_tc].[impven]))) AS imptotsol, IIf([com_compras].[imptot] Is Null,0,IIf([com_compras].[idmon]=2,[com_compras].[imptot],IIf([con_tc].[impven] Is Null,0,[com_compras].[imptot]/[con_tc].[impven]))) AS imptotdol, mae_documentocta.idcuen, con_planctas.cuenta as ctanum, con_planctas.descripcion as ctadesc " _
        + vbCr + " FROM ((mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (mae_prov RIGHT JOIN ((com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_prov.id = com_compras.idpro) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) INNER JOIN mae_documentocta ON (com_compras.idmon = mae_documentocta.idmon) AND (com_compras.tipdoc = mae_documentocta.iddoc)) INNER JOIN con_planctas ON mae_documentocta.idcuen = con_planctas.id " _
        + vbCr + " WHERE (((com_compras.fchdoc)<=CDate('" & TxtFecha.Valor & "')) AND ((com_compras.idmon)=2) AND ((com_compras.tipdoc)<>7) AND ((mae_documentocta.tipope)=0)) " _
        + vbCr + " ORDER BY com_compras.fchdoc, com_compras.numreg asc "

        
    RST_Busq rst, nSQL, xCon
    
    
    '--cargando la lista de pagos
    nSQL = "SELECT pago.rregistro, Last(pago.idmesope) AS idmes, Last(pago.fchope) AS fchope, Last(pago.tipcam) AS tipcam, Sum(pago.imptotal) AS tot, Sum(pago.imptotsol) AS totsol, Sum(pago.imptotdol) AS totdol, pago.iddoc " _
        + vbCr + " FROM (SELECT con_diario.rregistro, con_diario.idmes as idmesope, con_diario.fchdoc AS fchope, mae_moneda.simbolo, " _
        + vbCr + " iif(con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven)) AS tipcam, " _
        + vbCr + " IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS imptotal, " _
        + vbCr + " IIf(con_diario.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, " _
        + vbCr + " IIf(con_diario.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) AS imptotdol, " _
        + vbCr + " IIf(con_diario.idlib=6,tes_cajadestinodet.iddoc,con_diario.iddocpro) AS iddoc " _
        + vbCr + " FROM ((con_diario LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN tes_cajadestinodet ON (con_diario.idmov = tes_cajadestinodet.idtes) AND (con_diario.iddocpro = tes_cajadestinodet.corr) " _
        + vbCr + " WHERE (((con_diario.idlib) In (6,8,39)) AND ((con_diario.ridlib) in (1) )) " _
        + vbCr + " ) AS pago " _
        + vbCr + " GROUP BY pago.rregistro, pago.iddoc " _
        + vbCr + " ORDER BY Last(pago.idmesope), Last(pago.fchope), Sum(pago.imptotsol);"
    
    RST_Busq rstPago, nSQL, xCon
    
    
    Fg1.Rows = Fg1.FixedRows
    DoEvents
    
    Frame1.Visible = True
    ProgressBar2.Min = 0
    ProgressBar2.Value = 0
    ProgressBar2.Max = rst.RecordCount
    
    If rst.RecordCount <> 0 Then
    
        Do While Not rst.EOF
            DoEvents
            ProgressBar2.Value = ProgressBar2.Value + 1
            Fg1.Rows = Fg1.Rows + 1
        
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = rst("id")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = -1
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(rst("registro"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(rst("numruc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(rst("nombre"))
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(rst("abrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosC(rst("numdoc2"))
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosC(rst("fchdoc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosC(rst("simbolo"))
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosC(rst("tipcam"))
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosN(rst("imptotal"))
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosC(rst("glosa"))
            
            rstPago.Filter = "iddoc= " & rst("id")
            
            If rstPago.RecordCount <> 0 Then
            
                Fg1.TextMatrix(Fg1.Rows - 1, 13) = MonthName(rstPago("idmes"), True)
                Fg1.TextMatrix(Fg1.Rows - 1, 14) = NulosC(rstPago("fchope"))
                Fg1.TextMatrix(Fg1.Rows - 1, 15) = NulosN(rstPago("tipcam"))
                Fg1.TextMatrix(Fg1.Rows - 1, 16) = NulosN(rstPago("tot"))
                
                Fg1.TextMatrix(Fg1.Rows - 1, 17) = NulosN(rstPago("totsol")) '--debe sol
                Fg1.TextMatrix(Fg1.Rows - 1, 18) = NulosN(rst("imptotsol")) '--haber sol
                Fg1.TextMatrix(Fg1.Rows - 1, 19) = NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 18)) - NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 17))
                '--provicion - pago
                            
                Fg1.TextMatrix(Fg1.Rows - 1, 20) = NulosN(rstPago("totdol")) '--debe sol
                Fg1.TextMatrix(Fg1.Rows - 1, 21) = NulosN(rst("imptotdol")) '--haber sol
                Fg1.TextMatrix(Fg1.Rows - 1, 22) = NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 21)) - NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 20))
                
                Fg1.TextMatrix(Fg1.Rows - 1, 23) = NulosC(rst("ctanum"))
                Fg1.TextMatrix(Fg1.Rows - 1, 24) = NulosC(rst("ctadesc"))
                
                If Abs(NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 22))) > 1.5 Then
'                    Fg1.Rows = Fg1.Rows - 1
'                    GoTo Seguir:
                    '***************************************************************************************
                    '--desactivar el registro para que no se haga el ajuste correspondiente
                    Fg1.TextMatrix(Fg1.Rows - 1, 2) = 0
                    
                    Fg1.TextMatrix(Fg1.Rows - 1, 17) = NulosN(rst("imptotdol")) * rstPago("tipcam") '--debe sol
                    Fg1.TextMatrix(Fg1.Rows - 1, 18) = NulosN(rst("imptotsol")) '--haber sol
                    Fg1.TextMatrix(Fg1.Rows - 1, 19) = NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 18)) - NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 17))
                    
                                
'                    Fg1.TextMatrix(Fg1.Rows - 1, 20) = NulosN(rst("imptotdol")) '--debe sol
'                    Fg1.TextMatrix(Fg1.Rows - 1, 21) = NulosN(rst("imptotdol")) '--haber sol
'                    Fg1.TextMatrix(Fg1.Rows - 1, 22) = 0 'NulosN(rstPago("totdol")) - NulosN(rst("imptotdol"))
                    
                    If Abs(NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 19))) = 0 Then
                        Fg1.Rows = Fg1.Rows - 1
                        GoTo Seguir:
                    End If
                    
                    '--solo mostrar los documentos pendientes, los cancelado no se mostraran
                    If OptEstado(0).Value = True Then
                        Fg1.Rows = Fg1.Rows - 1
                        GoTo Seguir:
                    End If
                    
                    '--cambiar de color a las filas con pagos a cuenta
                    '--considerara pagado en su totalidad
                    GRID_COLOR_FONDO Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, &H9BFF79
                    
                    '***************************************************************************************
                ElseIf Abs(NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 19))) = 0 Then
                    Fg1.Rows = Fg1.Rows - 1
                    GoTo Seguir:
                
                Else
                    '--muestra los documentos cancelados o en su defecto documentos con saldo +- 1 .5
                
                    '--solo mostrar los documentos cancelados, los pendientes no se mostraran
                    If OptEstado(1).Value = True Then
                        Fg1.Rows = Fg1.Rows - 1
                        GoTo Seguir:
                    End If
                
                End If
                
                '--especificar si es ganancia o perdida
                If rst("idmon") = 2 Then
                    If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 19)) > 0 Then '--
                        Fg1.TextMatrix(Fg1.Rows - 1, 25) = "Si"
                        Fg1.TextMatrix(Fg1.Rows - 1, 26) = ""
                    Else
                        Fg1.TextMatrix(Fg1.Rows - 1, 25) = ""
                        Fg1.TextMatrix(Fg1.Rows - 1, 26) = "Si"
                    End If

                Else
                    If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 22)) > 0 Then '--
                        Fg1.TextMatrix(Fg1.Rows - 1, 25) = "Si"
                        Fg1.TextMatrix(Fg1.Rows - 1, 26) = ""
                    Else
                        Fg1.TextMatrix(Fg1.Rows - 1, 25) = ""
                        Fg1.TextMatrix(Fg1.Rows - 1, 26) = "Si"
                    End If
                
                End If
                
                
                
            Else
                '--eliminar fila si no tiene pagos
                Fg1.Rows = Fg1.Rows - 1
                GoTo Seguir:
            End If
            
            Fg1.TextMatrix(Fg1.Rows - 1, 27) = NulosN(rst("idcuen"))
            Fg1.TextMatrix(Fg1.Rows - 1, 28) = NulosN(rst("idpro"))
            Fg1.TextMatrix(Fg1.Rows - 1, 29) = NulosC(rst("tipdoc"))
            
            
Seguir:
            
            rst.MoveNext
        Loop
        
    
    End If
    Frame1.Visible = False
    Set rst = Nothing
    Set rstPago = Nothing


End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



Private Sub pConsultar()
''--Modificado 23/09/11 por Johan Castro
''       Agregar variable nSQLNC; Agregar las NC de compra y venta como pago para reducir el saldo final

    Dim nSQL As String
    Dim rst As New ADODB.Recordset
    Dim rstPago As New ADODB.Recordset
    
    Dim nSQLIdLib As String
    
    Dim sGan As Double '--variable que acumula ganancia por dif cambio
    Dim sPer As Double '--variable que acumula perdida por dif cambio
    
    Dim nSQLNC As String '--utilizado solamente para las notas de credito de compras y ventas
    
    lblGan.Caption = "0.00"
    lblPer.Caption = "0.00"
    sGan = 0
    sPer = 0
    Fg1.Rows = Fg1.FixedRows
    DoEvents
    '------------------------
    '--LblIdAjuste.Caption indica el ajuste en la moneda
    
    '--cargando datos de provicion
    '--cargando
    Select Case NulosN(LblIdLibro.Caption)
        Case 1 '--compras
            nSQL = "SELECT com_compras.id, com_compras.idmon,com_compras.idpro,com_compras.tipdoc,  com_compras.glosa,mae_prov.numruc, mae_prov!nombre AS nombre, IIf(com_compras.numreg Is Null Or com_compras.numreg='','',Left(com_compras.numreg,2) & mae_libros.codsun & Right(com_compras.numreg,4)) AS registro, " _
                + vbCr + " 'Compras' AS libro, mae_documento.abrev, IIf(com_compras!numser Is Null Or com_compras!numser='','',com_compras!numser+'-')+com_compras!numdoc AS numdoc2,  com_compras.fchdoc, mae_moneda.simbolo, iif(com_compras.tc=0, con_tc.impven,com_compras.tc) AS tipcam, " _
                + vbCr + " com_compras.imptot AS imptotal, IIf(com_compras.idmon=1,com_compras.imptot,com_compras.imptot * tipcam ) AS imptotsol, IIf(com_compras.idmon=2,com_compras.imptot,IIf(tipcam =0,0,com_compras.imptot/tipcam)) AS imptotdol, " _
                + vbCr + " mae_documentocta.idcuen, con_planctas.cuenta as ctanum, con_planctas.descripcion as ctadesc " _
                + vbCr + " FROM ((mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (mae_prov RIGHT JOIN ((com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) ON mae_prov.id = com_compras.idpro) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) INNER JOIN mae_documentocta ON (com_compras.idmon = mae_documentocta.idmon) AND (com_compras.tipdoc = mae_documentocta.iddoc)) INNER JOIN con_planctas ON mae_documentocta.idcuen = con_planctas.id " _
                + vbCr + " WHERE (((com_compras.fchdoc)<=CDate('" & TxtFecha.Valor & "')) AND ((com_compras.idmon)=" & NulosN(LblIdMon.Caption) & ") AND ((com_compras.tipdoc)<>7) AND ((mae_documentocta.tipope)=0)) " _
                + vbCr + " ORDER BY com_compras.fchdoc, com_compras.numreg asc "
            
            
            nSQLIdLib = "6,8,39"
            '--variable para identificar el proveedor en diario
            mIdTipPer = 1
            
nSQLNC = vbCr & " UNION " _
    + vbCr + " SELECT Left([com_compras].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([com_compras].[numreg],4) AS rregistro, " _
    + vbCr + " Month([com_compras].[fchreg]) AS idmesope, com_compras_1.fchdoc AS fchope, mae_moneda.simbolo, " _
    + vbCr + " IIf([com_compras_1].[tc]=0,[con_tc].[impven],[com_compras_1].[tc]) AS tipcam, com_compras_1.imptot AS imptotal, " _
    + vbCr + " IIf(com_compras_1.idmon=1,com_compras_1.imptot,com_compras_1.imptot*tipcam) AS imptotsol, IIf(com_compras_1.idmon=2,com_compras_1.imptot,IIf(tipcam=0,0,com_compras_1.imptot/tipcam)) AS imptotdol, " _
    + vbCr + " com_compras_1.iddocref AS iddoc " _
    + vbCr + " FROM ((com_compras AS com_compras_1 INNER JOIN (com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON com_compras_1.iddocref = com_compras.id) LEFT JOIN mae_moneda ON com_compras_1.idmon = mae_moneda.id) LEFT JOIN con_tc ON com_compras_1.fchdoc = con_tc.fecha " _
    + vbCr + " WHERE (((com_compras_1.tipdoc)=7)) "
            
            
            
            
        Case 2 '--ventas
            nSQL = "SELECT vta_ventas.id, vta_ventas.idmon, vta_ventas.idcli AS idpro, vta_ventas.tipdoc, vta_ventas.glosa, IIf(vta_ventas!anulado=-1,' ',mae_cliente!numruc) AS numruc, IIf(vta_ventas!anulado=-1,'Anulado',mae_cliente!nombre) AS nombre, IIf(vta_ventas.numreg Is Null Or vta_ventas.numreg='',mae_libros.codsun,Left(vta_ventas.numreg,2) & mae_libros.codsun & Right(vta_ventas.numreg,4)) AS registro, " _
                + vbCr + " 'Ventas' AS libro, mae_documento.abrev, vta_ventas!numser+'-'+vta_ventas!numdoc AS numdoc2, vta_ventas.fchdoc, mae_moneda.simbolo, IIf(vta_ventas.tc Is Null Or vta_ventas.tc=0,con_tc.impven,vta_ventas.tc) AS tipcam, " _
                + vbCr + " vta_ventas.imptotdoc AS imptotal, vta_ventas.impsal, IIf(imptotal=0,0,IIf(vta_ventas.idmon=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, IIf(imptotal=0,0,IIf(vta_ventas.idmon=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, vta_ventas.glosa AS glosaope, mae_documentocta.idcuen, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_documentocta.tipope " _
                + vbCr + " FROM con_planctas RIGHT JOIN ((((((vta_ventas LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN mae_documentocta ON (vta_ventas.idmon = mae_documentocta.idmon) AND (vta_ventas.tipdoc = mae_documentocta.iddoc)) ON con_planctas.id = mae_documentocta.idcuen " _
                + vbCr + " WHERE vta_ventas.idmon=" & NulosN(LblIdMon.Caption) & " and  (((vta_ventas.fchdoc)<=CDate('" & TxtFecha.Valor & "')) AND ((vta_ventas.tipdoc)<>7) AND ((vta_ventas.anulado)=0) AND ((mae_documentocta.tipope)=-1)) " _
                + vbCr + " ORDER BY IIf(vta_ventas!anulado=-1,'Anulado',mae_cliente!nombre), vta_ventas!numser+'-'+vta_ventas!numdoc;"
    
            nSQLIdLib = "5,6,8,37"
            '--variable para identificar el cliente en diario
            mIdTipPer = 2

nSQLNC = vbCr & " UNION " _
    + vbCr + " SELECT Left([vta_ventas].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([vta_ventas].[numreg],4) AS rregistro, " _
    + vbCr + " Month([vta_ventas].[fchreg]) AS idmesope, vta_ventas_1.fchdoc AS fchope, mae_moneda.simbolo, " _
    + vbCr + " IIf([vta_ventas_1].[tc]=0,[con_tc].[impven],[vta_ventas_1].[tc]) AS tipcam, vta_ventas_1.imptotdoc AS imptotal, " _
    + vbCr + " IIf(vta_ventas_1.idmon=1,vta_ventas_1.imptotdoc,vta_ventas_1.imptotdoc*tipcam) AS imptotsol, IIf(vta_ventas_1.idmon=2,vta_ventas_1.imptotdoc,IIf(tipcam=0,0,vta_ventas_1.imptotdoc/tipcam)) AS imptotdol, " _
    + vbCr + " vta_ventas_1.iddocref AS iddoc " _
    + vbCr + " FROM ((vta_ventas AS vta_ventas_1 INNER JOIN (vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) ON vta_ventas_1.iddocref = vta_ventas.id) LEFT JOIN mae_moneda ON vta_ventas_1.idmon = mae_moneda.id) LEFT JOIN con_tc ON vta_ventas_1.fchdoc = con_tc.fecha " _
    + vbCr + " WHERE (((vta_ventas_1.tipdoc)=7)) "

        Case 4 '--percepcion
                MsgBox "pendiente", vbInformation, xTitulo
                Exit Sub
                '--variable para identificar el proveedor en diario
                mIdTipPer = 1
        Case 37 '--letras
        
            nSQL = "SELECT let_letradet.corr AS id, let_letra.idmon, let_letra.idclipro AS idpro, let_letra.tipdoc, let_letra.glosa, mae_cliente.numruc, mae_cliente.nombre, Left([let_letra].[numreg],2) & [mae_libros].[codsun] & Right([let_letra].[numreg],4) AS registro, " _
                + vbCr + " 'Letras' AS libro, mae_documento.abrev, [let_letra].[ano] & ' ' & [let_letradet].[numdoc] & ' ' & [let_letradet].[numser] AS numdoc2, let_letradet.fchemi AS fchdoc, mae_moneda.simbolo, IIf([let_letra].[tc]=0,[con_tc].[impven],[let_letra].[tc]) AS tipcam, " _
                + vbCr + " let_letradet.implet AS imptotal, IIf(let_letra.idmon=1,let_letradet.implet,let_letradet.implet*tipcam) AS imptotsol, IIf(let_letra.idmon=2,let_letradet.implet,IIf(tipcam=0,0,let_letradet.implet/tipcam)) AS imptotdol, " _
                + vbCr + " mae_letra.idcuenven AS idcuen, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc " _
                + vbCr + " FROM (mae_moneda RIGHT JOIN (((((mae_cliente RIGHT JOIN let_letra ON mae_cliente.id = let_letra.idclipro) LEFT JOIN mae_documento ON let_letra.tipdoc = mae_documento.id) LEFT JOIN mae_libros ON let_letra.idlib = mae_libros.id) LEFT JOIN con_tc ON let_letra.fchemi = con_tc.fecha) INNER JOIN let_letradet ON let_letra.id = let_letradet.idlet) ON mae_moneda.id = let_letra.idmon) LEFT JOIN (mae_letra LEFT JOIN con_planctas ON mae_letra.idcuenven = con_planctas.id) ON let_letra.idmon = mae_letra.idmon " _
                + vbCr + " WHERE (((let_letradet.fchemi)<=CDate('" & TxtFecha.Valor & "')) AND ((let_letra.idmon)=" & NulosN(LblIdMon.Caption) & "))" _
                + vbCr + " ORDER BY mae_cliente.nombre, [let_letra].[ano] & ' ' & [let_letradet].[numdoc] & ' ' & [let_letradet].[numser];"
        
            nSQLIdLib = "6,42"
            '--variable para identificar el cliente en diario
            mIdTipPer = 2
            
        Case 40 '--honorarios
                nSQL = "SELECT com_honorarios.id, com_honorarios.idmon,com_honorarios.idpro,com_honorarios.tipdoc,  com_honorarios.glosa,mae_prov.numruc, mae_prov!nombre AS nombre, IIf(com_honorarios.numreg Is Null Or com_honorarios.numreg='','',Left(com_honorarios.numreg,2) & mae_libros.codsun & Right(com_honorarios.numreg,4)) AS registro, " _
                    + vbCr + " 'Honorarios' AS libro, mae_documento.abrev, IIf(com_honorarios!numser Is Null Or com_honorarios!numser='','',com_honorarios!numser+'-')+com_honorarios!numdoc AS numdoc2,  com_honorarios.fchdoc, mae_moneda.simbolo, iif(com_honorarios.tc=0, con_tc.impven,com_honorarios.tc) AS tipcam, " _
                    + vbCr + " com_honorarios.imptot AS imptotal, IIf(com_honorarios.idmon=1,com_honorarios.imptot,com_honorarios.imptot * tipcam ) AS imptotsol, IIf(com_honorarios.idmon=2,com_honorarios.imptot,IIf(tipcam =0,0,com_honorarios.imptot/tipcam)) AS imptotdol, " _
                    + vbCr + " mae_documentocta.idcuen, con_planctas.cuenta as ctanum, con_planctas.descripcion as ctadesc " _
                    + vbCr + " FROM ((mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (mae_prov RIGHT JOIN ((com_honorarios LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha) ON mae_prov.id = com_honorarios.idpro) ON mae_documento.id = com_honorarios.tipdoc) ON mae_moneda.id = com_honorarios.idmon) INNER JOIN mae_documentocta ON (com_honorarios.idmon = mae_documentocta.idmon) AND (com_honorarios.tipdoc = mae_documentocta.iddoc)) INNER JOIN con_planctas ON mae_documentocta.idcuen = con_planctas.id " _
                    + vbCr + " WHERE (((com_honorarios.fchdoc)<=CDate('" & TxtFecha.Valor & "')) AND ((com_honorarios.idmon)=" & NulosN(LblIdMon.Caption) & ") AND ((com_honorarios.tipdoc)<>7) AND ((mae_documentocta.tipope)=0)) " _
                    + vbCr + " ORDER BY com_honorarios.fchdoc, com_honorarios.numreg asc "
                    
             nSQLIdLib = "6,39"
        
            '--variable para identificar el proveedor en diario
            mIdTipPer = 1
        
        Case 41 '--lgd
            nSQL = "SELECT vta_gastodebito.id, vta_gastodebito.idmon, vta_gastodebito.idcli AS idpro, vta_gastodebito.tipdoc, vta_gastodebito.glosa, IIf([vta_gastodebito]![anulado]=-1,' ',[mae_cliente]![numruc]) AS numruc, IIf([vta_gastodebito]![anulado]=-1,'Anulado',[mae_cliente]![nombre]) AS nombre, IIf([vta_gastodebito].[numreg] Is Null Or [vta_gastodebito].[numreg]='',[mae_libros].[codsun],Left([vta_gastodebito].[numreg],2) & [mae_libros].[codsun] & Right([vta_gastodebito].[numreg],4)) AS registro, " _
                    + vbCr + " 'Lgd' AS libro, mae_documento.abrev, [vta_gastodebito]![numser]+'-'+[vta_gastodebito]![numdoc] AS numdoc2, vta_gastodebito.fchemi AS fchdoc, mae_moneda.simbolo, IIf(vta_gastodebito.tc Is Null Or vta_gastodebito.tc=0,con_tc.impven,vta_gastodebito.tc) AS tipcam, " _
                    + vbCr + " vta_gastodebito.imptot AS imptotal, IIf([vta_gastodebito].[idmon]=1,imptotal,imptotal*tipcam) AS imptotsol, IIf([vta_gastodebito].[idmon]=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) AS imptotdol, " _
                    + vbCr + " mae_documentocta.idcuen, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc " _
                    + vbCr + " FROM (((((vta_gastodebito LEFT JOIN mae_cliente ON vta_gastodebito.idcli = mae_cliente.id) LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id) LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha) LEFT JOIN mae_libros ON vta_gastodebito.idlib = mae_libros.id) LEFT JOIN mae_moneda ON vta_gastodebito.idmon = mae_moneda.id) LEFT JOIN (con_planctas RIGHT JOIN mae_documentocta ON con_planctas.id = mae_documentocta.idcuen) ON (vta_gastodebito.idmon = mae_documentocta.idmon) AND (vta_gastodebito.tipdoc = mae_documentocta.iddoc) " _
                    + vbCr + " WHERE (((vta_gastodebito.fchemi)<=CDate('" & TxtFecha.Valor & "')) AND ((vta_gastodebito.idmon)=" & NulosN(LblIdMon.Caption) & ") AND ((mae_documentocta.tipope)=-1) AND ((vta_gastodebito.anulado)=0) AND ((vta_gastodebito.tipdoc)<>126)) " _
                    + vbCr + " ORDER BY IIf([vta_gastodebito]![anulado]=-1,'Anulado',[mae_cliente]![nombre]), [vta_gastodebito]![numser]+'-'+[vta_gastodebito]![numdoc]; "

            '--variable para identificar el cliente en diario
            mIdTipPer = 1
            
            nSQLIdLib = "6"
        Case 42 '--planilla de letras
            nSQL = "SELECT let_planilla.id, let_planilla.idmon, mae_banconumcta.idban AS idpro, let_planilla.tipdoc, let_planilla.glosa, mae_bancos.numruc, mae_bancos.descripcion AS nombre, Left([let_planilla].[numreg],2) & [mae_libros].[codsun] & Right([let_planilla].[numreg],4) AS registro, " _
                    + vbCr + " 'Planilla letra' AS libro, mae_documento.abrev, let_planilla.numdoc AS numdoc2, let_planilla.fchemi AS fchdoc, mae_moneda.simbolo, IIf([let_planilla].[anulado]=-1,0,IIf([let_planilla].[tc]=0,[con_tc].[impven],[let_planilla].[tc])) AS tipcam,  " _
                    + vbCr + " let_planilla.imptot AS imptotal, IIf(let_planilla.idmon=1,let_planilla.imptot,let_planilla.imptot*tipcam) AS imptotsol, IIf(let_planilla.idmon=2,let_planilla.imptot,IIf(tipcam=0,0,let_planilla.imptot/tipcam)) AS imptotdol, " _
                    + vbCr + " let_modalidadctabco.idcuen,  con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc " _
                    + vbCr + " FROM ((mae_documento RIGHT JOIN (mae_bancos RIGHT JOIN ((((let_planilla LEFT JOIN mae_moneda ON let_planilla.idmon = mae_moneda.id) LEFT JOIN mae_banconumcta ON let_planilla.idbcocta = mae_banconumcta.id) LEFT JOIN mae_libros ON let_planilla.idlib = mae_libros.id) LEFT JOIN con_tc ON let_planilla.fchemi = con_tc.fecha) ON mae_bancos.id = mae_banconumcta.idban) ON mae_documento.id = let_planilla.tipdoc) LEFT JOIN let_modalidadctabco ON (let_planilla.idmon = let_modalidadctabco.idmon) AND (let_planilla.idbcocta = let_modalidadctabco.idbcocta) AND (let_planilla.idmod = let_modalidadctabco.idmod)) LEFT JOIN con_planctas ON let_modalidadctabco.idcuen = con_planctas.id " _
                    + vbCr + " WHERE let_planilla.fchemi<=CDate('" & TxtFecha.Valor & "') and  let_planilla.idmon=" & NulosN(LblIdMon.Caption) & "  " _
                    + vbCr + " ORDER BY mae_bancos.descripcion, let_planilla.numdoc;"
                
            nSQLIdLib = "6"
            
            '--variable para identificar el proveedor en diario
            mIdTipPer = 1
        Case Else
            
            Exit Sub
        End Select
    
        
        
    RST_Busq rst, nSQL, xCon
    
        
    '--cargando la lista de pagos
    nSQL = "SELECT pago.rregistro, Last(pago.idmesope) AS idmes, Last(pago.fchope) AS fchope, Last(pago.tipcam) AS tipcam, Sum(pago.imptotal) AS tot, Sum(pago.imptotsol) AS totsol, Sum(pago.imptotdol) AS totdol, pago.iddoc " _
        + vbCr + " FROM (SELECT con_diario.rregistro, con_diario.idmes as idmesope, con_diario.fchdoc AS fchope, mae_moneda.simbolo, " _
        + vbCr + " iif(con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven)) AS tipcam, " _
        + vbCr + " IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS imptotal, " _
        + vbCr + " IIf(con_diario.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, " _
        + vbCr + " IIf(con_diario.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) AS imptotdol, " _
        + vbCr + " con_diario.iddoc " _
        + vbCr + " FROM (con_diario LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
        + vbCr + " WHERE (((con_diario.idlib) In (" & nSQLIdLib & ")) AND ((con_diario.ridlib) in (" & NulosN(LblIdLibro.Caption) & ") )) " & nSQLNC _
        + vbCr + " ) AS pago " _
        + vbCr + " GROUP BY pago.rregistro, pago.iddoc " _
        + vbCr + " ORDER BY Last(pago.idmesope), Last(pago.fchope), Sum(pago.imptotsol);"
    
    RST_Busq rstPago, nSQL, xCon
    
    
    Fg1.Rows = Fg1.FixedRows
    DoEvents
    
    
    
    If rst.RecordCount <> 0 Then
                
        '--mostrando la barra de progreso
        Frame1.Visible = True
        ProgressBar2.Min = 0
        ProgressBar2.Value = 0
        '--asignar cantidad de registros encontrados
        ProgressBar2.Max = rst.RecordCount
        
        Do While Not rst.EOF
            DoEvents
            ProgressBar2.Value = ProgressBar2.Value + 1
            Fg1.Rows = Fg1.Rows + 1
        
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = rst("id")
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = -1
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(rst("registro"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(rst("numruc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(rst("nombre"))
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosC(rst("abrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosC(rst("numdoc2"))
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosC(rst("fchdoc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosC(rst("simbolo"))
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosC(rst("tipcam"))
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(NulosN(rst("imptotal")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosC(rst("glosa"))
            
            rstPago.Filter = "iddoc= " & rst("id")
            
            If rstPago.RecordCount <> 0 Then
            
                Fg1.TextMatrix(Fg1.Rows - 1, 13) = MonthName(rstPago("idmes"), True)
                Fg1.TextMatrix(Fg1.Rows - 1, 14) = NulosC(rstPago("fchope"))
                Fg1.TextMatrix(Fg1.Rows - 1, 15) = NulosN(rstPago("tipcam"))
                Fg1.TextMatrix(Fg1.Rows - 1, 16) = NulosN(rstPago("tot"))
                
                Fg1.TextMatrix(Fg1.Rows - 1, 17) = Format(NulosN(rstPago("totsol")), FORMAT_MONTO) '--debe sol
                Fg1.TextMatrix(Fg1.Rows - 1, 18) = Format(NulosN(rst("imptotsol")), FORMAT_MONTO)  '--haber sol
                Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 18)) - NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 17)), FORMAT_MONTO)
                '--provicion - (pago/cobranza)
                            
                Fg1.TextMatrix(Fg1.Rows - 1, 20) = Format(NulosN(rstPago("totdol")), FORMAT_MONTO)  '--debe sol
                Fg1.TextMatrix(Fg1.Rows - 1, 21) = Format(NulosN(rst("imptotdol")), FORMAT_MONTO)   '--haber sol
                Fg1.TextMatrix(Fg1.Rows - 1, 22) = Format(NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 21)) - NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 20)), FORMAT_MONTO)
                
                Fg1.TextMatrix(Fg1.Rows - 1, 23) = NulosC(rst("ctanum"))
                Fg1.TextMatrix(Fg1.Rows - 1, 24) = NulosC(rst("ctadesc"))
                
                If Abs(NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 22))) > 1.5 Then
'                    Fg1.Rows = Fg1.Rows - 1
'                    GoTo Seguir:
                    '***************************************************************************************
                    '--desactivar el registro para que no se haga el ajuste correspondiente
                    Fg1.TextMatrix(Fg1.Rows - 1, 2) = 0
                    
                    Fg1.TextMatrix(Fg1.Rows - 1, 17) = Format(NulosN(rst("imptotdol")) * rstPago("tipcam"), FORMAT_MONTO) '--debe sol
                    Fg1.TextMatrix(Fg1.Rows - 1, 18) = Format(NulosN(rst("imptotsol")), FORMAT_MONTO) '--haber sol
                    Fg1.TextMatrix(Fg1.Rows - 1, 19) = Format(NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 18)) - NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 17)), FORMAT_MONTO)
                    
                                
'                    Fg1.TextMatrix(Fg1.Rows - 1, 20) = NulosN(rst("imptotdol")) '--debe sol
'                    Fg1.TextMatrix(Fg1.Rows - 1, 21) = NulosN(rst("imptotdol")) '--haber sol
'                    Fg1.TextMatrix(Fg1.Rows - 1, 22) = 0 'NulosN(rstPago("totdol")) - NulosN(rst("imptotdol"))
                    
                    If Abs(NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 19))) = 0 Then
                        Fg1.Rows = Fg1.Rows - 1
                        GoTo Seguir:
                    End If
                    
                    '--solo mostrar los documentos pendientes, los cancelado no se mostraran
                    If OptEstado(0).Value = True Then
                        Fg1.Rows = Fg1.Rows - 1
                        GoTo Seguir:
                    End If
                    
                    '--cambiar de color a las filas con pagos a cuenta
                    '--considerara pagado en su totalidad
                    GRID_COLOR_FONDO Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, Fg1.Cols - 1, &H9BFF79
                    
                    '***************************************************************************************
                ElseIf Abs(NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 19))) = 0 Then
                    Fg1.Rows = Fg1.Rows - 1
                    GoTo Seguir:
                
                Else
                    '--muestra los documentos cancelados o en su defecto documentos con saldo +- 1 .5
                
                    '--solo mostrar los documentos cancelados, los pendientes no se mostraran
                    If OptEstado(1).Value = True Then
                        Fg1.Rows = Fg1.Rows - 1
                        GoTo Seguir:
                    End If
                
                End If
                
                '--especificar si es ganancia o perdida
                If rst("idmon") = 2 Then
                    If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 19)) > 0 Then '--
                        Fg1.TextMatrix(Fg1.Rows - 1, 25) = "Si"
                        Fg1.TextMatrix(Fg1.Rows - 1, 26) = ""
                        sGan = sGan + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 19))
                    Else
                        Fg1.TextMatrix(Fg1.Rows - 1, 25) = ""
                        Fg1.TextMatrix(Fg1.Rows - 1, 26) = "Si"
                        sPer = sPer + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 19))
                    End If

                Else
                    If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 22)) > 0 Then '--
                        Fg1.TextMatrix(Fg1.Rows - 1, 25) = "Si"
                        Fg1.TextMatrix(Fg1.Rows - 1, 26) = ""
                        sGan = sGan + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 22))
                    Else
                        Fg1.TextMatrix(Fg1.Rows - 1, 25) = ""
                        Fg1.TextMatrix(Fg1.Rows - 1, 26) = "Si"
                        sPer = sPer + NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 22))
                    End If
                
                End If
                
                
                
            Else
                '--eliminar fila si no tiene pagos
                Fg1.Rows = Fg1.Rows - 1
                GoTo Seguir:
            End If
            
            Fg1.TextMatrix(Fg1.Rows - 1, 27) = NulosN(rst("idcuen"))
            Fg1.TextMatrix(Fg1.Rows - 1, 28) = NulosN(rst("idpro"))
            Fg1.TextMatrix(Fg1.Rows - 1, 29) = NulosC(rst("tipdoc"))
            
            
Seguir:
            
            rst.MoveNext
        Loop
        
    
    End If
    
    '--mostra resumen
    lblGan.Caption = Format(sGan, FORMAT_MONTO)
    lblPer.Caption = Format(sPer, FORMAT_MONTO)
    
    
    Frame1.Visible = False
    Set rst = Nothing
    Set rstPago = Nothing


End Sub


Private Sub CmdBusGan_Click()

    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Nº Cuenta":    xCampos(0, 1) = "cuenta":        xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "5000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id FROM con_planctas ORDER BY cuenta"
    
    xform.Titulo = "Buscando Cuenta Contable"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "cuenta"
    xform.CampoBusca = "cuenta"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        TxtGanancia.Text = xRs("cuenta")
        LblDescGanancia.Caption = xRs("descripcion")
        LblIdCtaGan.Caption = xRs("id")
        TxtPerdida.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub TxtGanancia_Change()
    If TxtGanancia.Text = "" Then
        Me.LblDescGanancia.Caption = ""
        Me.LblIdCtaGan.Caption = ""
    End If
End Sub

Private Sub CmdBusPer_Click()


    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Nº Cuenta":    xCampos(0, 1) = "cuenta":        xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "5000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id FROM con_planctas ORDER BY cuenta"
    
    xform.Titulo = "Buscando Cuenta Contable"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "cuenta"
    xform.CampoBusca = "cuenta"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        TxtPerdida.Text = xRs("cuenta")
        LblDescPerdida.Caption = xRs("descripcion")
        LblIdCtaPer.Caption = xRs("id")
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub


Private Sub TxtPerdida_Change()
    If TxtPerdida.Text = "" Then
        Me.LblDescPerdida.Caption = ""
        Me.LblIdCtaPer.Caption = ""
    End If
End Sub
