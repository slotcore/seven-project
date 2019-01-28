VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmRedondeoCentimos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Herramientas - Redondeo a Céntimos"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Resumen"
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
      Left            =   6270
      TabIndex        =   12
      Top             =   360
      Width           =   5445
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
         Left            =   870
         TabIndex        =   18
         Top             =   210
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tot. Reg:"
         Height          =   195
         Left            =   60
         TabIndex        =   17
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ganancia:"
         Height          =   195
         Left            =   1740
         TabIndex        =   16
         Top             =   330
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pérdida:"
         Height          =   195
         Left            =   3690
         TabIndex        =   15
         Top             =   330
         Width           =   585
      End
      Begin VB.Label lblGan 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblGan"
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
         Left            =   2550
         TabIndex        =   14
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label lblPer 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblPer"
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
         Left            =   4350
         TabIndex        =   13
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   735
      Left            =   3180
      TabIndex        =   5
      Top             =   2580
      Visible         =   0   'False
      Width           =   5805
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   285
         Left            =   90
         TabIndex        =   6
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
         TabIndex        =   8
         Top             =   135
         Width           =   1500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1695
         TabIndex        =   7
         Top             =   150
         Width           =   45
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid fg1 
      Height          =   5580
      Left            =   30
      TabIndex        =   4
      Top             =   990
      Width           =   11670
      _cx             =   20585
      _cy             =   9842
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
      FormatString    =   $"FrmRedondeoCentimos.frx":0000
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
      Left            =   8250
      Top             =   0
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
            Picture         =   "FrmRedondeoCentimos.frx":00D5
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRedondeoCentimos.frx":0619
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRedondeoCentimos.frx":09AB
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRedondeoCentimos.frx":0B2F
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRedondeoCentimos.frx":0F83
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRedondeoCentimos.frx":109B
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRedondeoCentimos.frx":15DF
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRedondeoCentimos.frx":1B23
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRedondeoCentimos.frx":1C37
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRedondeoCentimos.frx":1D4B
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRedondeoCentimos.frx":219F
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRedondeoCentimos.frx":230B
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRedondeoCentimos.frx":2853
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   9
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
      Height          =   630
      Left            =   30
      TabIndex        =   0
      Top             =   360
      Width           =   6225
      Begin VB.CommandButton CmdBusProv 
         Height          =   230
         Left            =   5340
         Picture         =   "FrmRedondeoCentimos.frx":2BE5
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   270
         Width           =   210
      End
      Begin VB.TextBox TxtRedondeo 
         Height          =   300
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "TxtRedondeo"
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label LblIdLibro 
         AutoSize        =   -1  'True
         Caption         =   "LblIdLibro"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   5010
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label LblIdMon 
         AutoSize        =   -1  'True
         Caption         =   "LblIdMon"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4860
         TabIndex        =   10
         Top             =   210
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label LblIdRedondeo 
         AutoSize        =   -1  'True
         Caption         =   "LblIdRedondeo"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2070
         TabIndex        =   3
         Top             =   90
         Visible         =   0   'False
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmRedondeoCentimos"
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



Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    TxtRedondeo.Text = ""
    
    lblGan.Caption = "0.00"
    lblPer.Caption = "0.00"
    lblTotReg.Caption = "0"
    
    Configurar_Grilla
    SeEjecuto = True
    
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
        .Cols = 15
        
        .ColWidth(0) = 200
        '--DATOS DE FILA
        
        SGI_JC1.GRID_COMBINAR Fg1, 0, 3, 0, 8, "DATOS DE LA OPERACIÓN", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC
        SGI_JC1.GRID_COMBINAR Fg1, 0, 9, 0, 11, "IMPORTE", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC
        SGI_JC1.GRID_COMBINAR Fg1, 0, 12, 0, 13, "P/G", flexAlignCenterCenter, True, , vbBlack, &HD8E9EC
      
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
        
        .TextMatrix(1, 12) = "Per":         .ColWidth(12) = 400:  .ColAlignment(12) = flexAlignCenterCenter:  .Row = 1: .Col = 12: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 13) = "Gan":         .ColWidth(13) = 400:  .ColAlignment(13) = flexAlignCenterCenter:  .Row = 1: .Col = 13: .CellAlignment = flexAlignCenterCenter
              
        .TextMatrix(1, 14) = "idmes":
        
        OCULTAR_COL Fg1, 14, 14
        
        Fg1.ColFormat(7) = "0.000"
        
        Fg1.ColFormat(5) = SGI_JC1.FORMAT_DATE
        
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
    
    nSQL = "SELECT mae_redondeo.*, per.cuenta AS ctagnum, per.descripcion AS ctagdesc, gan.cuenta AS ctapnum, gan.descripcion AS ctapdesc, mae_moneda.simbolo " _
            + vbCr + " FROM ((mae_redondeo LEFT JOIN mae_moneda ON mae_redondeo.idmon = mae_moneda.id) INNER JOIN con_planctas AS per ON mae_redondeo.idcuenper = per.id) INNER JOIN con_planctas AS gan ON mae_redondeo.idcuengan = gan.id;"
    
    xObj.CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Redondeo a Centimos", "descripcion", "descripcion", Principio
    Set xObj = Nothing
    If xRs.State = 1 Then
        TxtRedondeo.Text = NulosC(xRs("descripcion"))
        LblIdRedondeo.Caption = NulosC(xRs("id"))
        
        LblIdMon.Caption = NulosN(xRs("idmon"))
        LblIdLibro.Caption = NulosN(xRs("idlib"))

    End If
    
    Set xRs = Nothing
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 2 Then Grabar
    If Button.Index = 4 Then pExportar
'    If Button.Index = 4 Then pImprimir
    If Button.Index = 6 Then
        Unload Me
    End If
End Sub

Private Sub TxtRedondeo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtRedondeo_KeyUp(KeyCode As Integer, Shift As Integer)
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
    GRID_EXPORTAR_MSEXCELTMP Fg1, xCon, flexFileCustomText, True, "ORIGEN PARA DIFERENCIA DE CAMBIO", "DEL 01/01/08 AL 31/12/08", " "

    
'    Set xFun = Nothing
    
End Sub


Private Sub pImprimir()

    Dim xPrint As New SGI2_funciones.formularios
    
    Me.MousePointer = vbHourglass
    xPrint.Imprimir_x_VSFlexGrid Fg1, "CONSULTA DE ASIENTO Nº. " & NumRegistro, " ", "", False, True
    Set xPrint = Nothing
    Me.MousePointer = vbDefault
  
    
End Sub


Private Sub pConsultar()
    Dim nSQL As String
    Dim sGan, sPer As Double
    Dim rst As New ADODB.Recordset
        
    If NulosN(LblIdRedondeo.Caption) = 0 Then
        MsgBox "Seleccione el Redondeo del Modulo", vbExclamation, xTitulo
        TxtRedondeo.SetFocus
        Exit Sub
    End If
    
    If AnoTra = "" Then
        MsgBox "Año de Trabajo incorrecto", vbInformation, xTitulo
        Exit Sub
    End If
    
    lblGan.Caption = "0.00"
    lblPer.Caption = "0.00"
    
    lblTotReg.Caption = "0"
    
    sGan = 0
    sPer = 0
    
    DoEvents
    Me.MousePointer = vbHourglass
    '--cargando la lista de pagos
'''
'    --eliminamos los registros anteriormente grabados
'    xCon.Execute "delete from con_diario where idlib = " & NulosN(LblIdLibro.Caption) & " and idmon = " & NulosC(LblIdMon.Caption) & " and ridtipper=999"
'''    '------------------------------------------
        
    
    'En ME
    If LblIdLibro.Caption = 6 Then
        If NulosN(LblIdMon.Caption) = 2 Then
            nSQL = "SELECT banco.idlib, banco.idmes, banco.idmov, banco.idmon, banco.registro, banco.tipo1,banco.fchope, banco.glosa, Last(banco.tipcam) AS tc, banco.simbolo, Sum(banco.impdebesol) AS totdebsol, Sum(banco.imphabersol) AS tothabsol, Sum(banco.impdebedol) AS totdeb, Sum(banco.imphaberdol) AS tothab, Sum(banco.impdebesol)-Sum(banco.imphabersol) AS salsol, Sum(banco.impdebedol)-Sum(banco.imphaberdol) AS totsal " _
                + vbCr + " FROM ( " _
                + vbCr + " SELECT con_diario.idlib, con_diario.idmes, con_diario.idmov, con_diario.idmon, Format(con_diario.idmes,'00') & IIf(mae_libros.codsun Is Null,'',mae_libros.codsun) & con_diario.numasi AS registro, mae_tipomov.descripcion AS tipo1, mae_libros.descripcion AS libdesc, con_diario.fchdoc AS fchope, trim(con_diario.rglosaope) as glosa, " _
                + vbCr + " iif(con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven)) AS tipcam, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo, " _
                + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.impdebdol*tipcam),con_diario.impdebsol) AS impdebesol, " _
                + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.imphabdol*tipcam),con_diario.imphabsol) AS imphabersol, " _
                + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(tipcam=0 Or con_diario.impdebsol=0,0,(con_diario.impdebsol/tipcam))) AS impdebedol, " _
                + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(tipcam=0 Or con_diario.imphabsol=0,0,(con_diario.imphabsol/tipcam))) AS imphaberdol " _
                + vbCr + " FROM (((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) INNER JOIN tes_caja ON con_diario.idmov = tes_caja.id) INNER JOIN mae_tipomov ON tes_caja.tipmov = mae_tipomov.id " _
                + vbCr + " WHERE (((con_diario.idlib)=" & NulosN(LblIdLibro.Caption) & ") AND ((con_diario.ajuste) In (0,2)) AND ((con_diario.fchasi) Between CDate('01/01/" & AnoTra & "') And CDate('31/12/" & AnoTra & "'))) " _
                + vbCr + " ) AS banco " _
                + vbCr + " GROUP BY banco.idlib, banco.idmes, banco.idmov, banco.idmon,banco.tipo1, banco.registro, banco.fchope, banco.glosa, banco.simbolo " _
                + vbCr + " HAVING (((banco.idmon)=2) AND ((Sum(banco.impdebedol)-Sum(banco.imphaberdol)) Between 0.00001 And 5 Or (Sum(banco.impdebedol)-Sum(banco.imphaberdol)) Between -5 And -0.00001)) ORDER BY banco.registro ASC "
        Else
            '--En MN
            nSQL = "SELECT banco.idlib, banco.idmes, banco.idmov, banco.idmon, banco.registro, banco.tipo1,banco.fchope, banco.glosa, Last(banco.tipcam) AS tc, banco.simbolo, Sum(banco.impdebesol) AS totdeb, Sum(banco.imphabersol) AS tothab, Sum(banco.impdebedol) AS totdebdol, Sum(banco.imphaberdol) AS tothabdol, Sum(banco.impdebesol)-Sum(banco.imphabersol) AS totsal, Sum(banco.impdebedol)-Sum(banco.imphaberdol) AS saldol " _
                + vbCr + " FROM ( " _
                + vbCr + " SELECT con_diario.idlib, con_diario.idmes, con_diario.idmov, con_diario.idmon, Format(con_diario.idmes,'00') & IIf(mae_libros.codsun Is Null,'',mae_libros.codsun) & con_diario.numasi AS registro, mae_tipomov.descripcion AS tipo1, mae_libros.descripcion AS libdesc, con_diario.fchdoc AS fchope, trim(con_diario.rglosaope) as glosa, " _
                + vbCr + " iif(con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven)) AS tipcam, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo, " _
                + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.impdebdol*tipcam),con_diario.impdebsol) AS impdebesol, " _
                + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.imphabdol*tipcam),con_diario.imphabsol) AS imphabersol, " _
                + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(tipcam=0 Or con_diario.impdebsol=0,0,(con_diario.impdebsol/tipcam))) AS impdebedol, " _
                + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(tipcam=0 Or con_diario.imphabsol=0,0,(con_diario.imphabsol/tipcam))) AS imphaberdol " _
                + vbCr + " FROM (((((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) INNER JOIN tes_caja ON con_diario.idmov = tes_caja.id) INNER JOIN mae_tipomov ON tes_caja.tipmov = mae_tipomov.id " _
                + vbCr + " WHERE (((con_diario.idlib)=" & NulosN(LblIdLibro.Caption) & ") AND ((con_diario.ajuste) In (0,1)) AND ((con_diario.fchasi) Between CDate('01/01/" & AnoTra & "') And CDate('31/12/" & AnoTra & "'))) " _
                + vbCr + " ) AS banco " _
                + vbCr + " GROUP BY banco.idlib, banco.idmes, banco.idmov, banco.idmon,banco.tipo1, banco.registro, banco.fchope, banco.glosa, banco.simbolo " _
                + vbCr + " HAVING (((banco.idmon)=1) AND ((Sum(banco.impdebesol)-Sum(banco.imphabersol)) Between 0.0001 And 5 Or (Sum(banco.impdebesol)-Sum(banco.imphabersol)) Between -5 And -0.0001)) ORDER BY banco.registro ASC "
    
        End If
    Else
        
        If NulosN(LblIdMon.Caption) = 2 Then
            nSQL = "SELECT banco.idlib, banco.idmes, banco.idmov, banco.idmon, banco.registro, '' as tipo1,banco.fchope, banco.glosa, Last(banco.tipcam) AS tc, banco.simbolo, Sum(banco.impdebesol) AS totdebsol, Sum(banco.imphabersol) AS tothabsol, Sum(banco.impdebedol) AS totdeb, Sum(banco.imphaberdol) AS tothab, Sum(banco.impdebesol)-Sum(banco.imphabersol) AS salsol, Sum(banco.impdebedol)-Sum(banco.imphaberdol) AS totsal " _
                + vbCr + " FROM ( " _
                + vbCr + " SELECT con_diario.idlib, con_diario.idmes, con_diario.idmov, con_diario.idmon, Format(con_diario.idmes,'00') & IIf(mae_libros.codsun Is Null,'',mae_libros.codsun) & con_diario.numasi AS registro, '' AS tipo1, mae_libros.descripcion AS libdesc, con_diario.fchdoc AS fchope, con_diario.rglosaope & '' as glosa, " _
                + vbCr + " iif(con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven)) AS tipcam, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo, " _
                + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.impdebdol*tipcam),con_diario.impdebsol) AS impdebesol, " _
                + vbCr + " IIf(con_diario.idmon=2,IIf(tipcam=0,0,con_diario.imphabdol*tipcam),con_diario.imphabsol) AS imphabersol, " _
                + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(tipcam=0 Or con_diario.impdebsol=0,0,(con_diario.impdebsol/tipcam))) AS impdebedol, " _
                + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(tipcam=0 Or con_diario.imphabsol=0,0,(con_diario.imphabsol/tipcam))) AS imphaberdol " _
                + vbCr + " FROM (((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
                + vbCr + " WHERE con_diario.idmon = " & NulosN(LblIdMon.Caption) & " and  (((con_diario.idlib)=" & NulosN(LblIdLibro.Caption) & ") AND ((con_diario.ajuste) In (0,2)) AND ((con_diario.fchasi) Between CDate('01/01/" & AnoTra & "') And CDate('31/12/" & AnoTra & "'))) " _
                + vbCr + " ) AS banco " _
                + vbCr + " GROUP BY banco.idlib, banco.idmes, banco.idmov, banco.idmon,banco.tipo1, banco.registro, banco.fchope, banco.glosa, banco.simbolo " _
                + vbCr + " HAVING (((Sum(banco.impdebedol)-Sum(banco.imphaberdol)) Between 0.00001 And 5 Or (Sum(banco.impdebedol)-Sum(banco.imphaberdol)) Between -5 And -0.00001)) ORDER BY banco.registro ASC "
        Else
            '--En MN
            nSQL = "SELECT banco.idlib, banco.idmes, banco.idmov, banco.idmon, banco.registro, '' as tipo1,banco.fchope, banco.glosa, Last(banco.tipcam) AS tc, banco.simbolo, Sum(banco.impdebesol) AS totdeb, Sum(banco.imphabersol) AS tothab, Sum(banco.impdebedol) AS totdebdol, Sum(banco.imphaberdol) AS tothabdol, Sum(banco.impdebesol)-Sum(banco.imphabersol) AS totsal, Sum(banco.impdebedol)-Sum(banco.imphaberdol) AS saldol " _
                + vbCr + " FROM ( " _
                + vbCr + " SELECT con_diario.idlib, con_diario.idmes, con_diario.idmov, con_diario.idmon, Format(con_diario.idmes,'00') & IIf(mae_libros.codsun Is Null,'',mae_libros.codsun) & con_diario.numasi AS registro, '' AS tipo1, mae_libros.descripcion AS libdesc, con_diario.fchdoc AS fchope, con_diario.rglosaope & '' as glosa, " _
                + vbCr + " iif(con_diario.aplicatc=-1,con_diario.tc,iif(con_tc.impven is null,0,con_tc.impven)) AS tipcam, con_planctas.cuenta AS ctanum, con_planctas.descripcion AS ctadesc, mae_moneda.simbolo, " _
                + vbCr + " IIf(con_diario.idmon=2,IIf(tc=0,0,con_diario.impdebdol*tc),con_diario.impdebsol) AS impdebesol, " _
                + vbCr + " IIf(con_diario.idmon=2,IIf(tc=0,0,con_diario.imphabdol*tc),con_diario.imphabsol) AS imphabersol, " _
                + vbCr + " IIf(con_diario.idmon=2,con_diario.impdebdol,IIf(tipcam=0 Or con_diario.impdebsol=0,0,(con_diario.impdebsol/tipcam))) AS impdebedol, " _
                + vbCr + " IIf(con_diario.idmon=2,con_diario.imphabdol,IIf(tipcam=0 Or con_diario.imphabsol=0,0,(con_diario.imphabsol/tipcam))) AS imphaberdol " _
                + vbCr + " FROM (((con_planctas RIGHT JOIN con_diario ON con_planctas.id = con_diario.idcue) LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha " _
                + vbCr + " WHERE con_diario.idmon = " & NulosN(LblIdMon.Caption) & " and (((con_diario.idlib)=" & NulosN(LblIdLibro.Caption) & ") AND ((con_diario.ajuste) In (0,1)) AND ((con_diario.fchasi) Between CDate('01/01/" & AnoTra & "') And CDate('31/12/" & AnoTra & "'))) " _
                + vbCr + " ) AS banco " _
                + vbCr + " GROUP BY banco.idlib, banco.idmes, banco.idmov, banco.idmon,banco.tipo1, banco.registro, banco.fchope, banco.glosa, banco.simbolo " _
                + vbCr + " HAVING ( ((Sum(banco.impdebesol)-Sum(banco.imphabersol)) Between 0.00001 And 5 Or (Sum(banco.impdebesol)-Sum(banco.imphabersol)) Between -5 And -0.00001)) ORDER BY banco.registro ASC "
    
        End If
        
    End If
    
    Fg1.Rows = Fg1.FixedRows
    DoEvents
    
    
    RST_Busq rst, nSQL, xCon
    DoEvents
    
    Frame1.Visible = True
    ProgressBar2.Min = 0
    ProgressBar2.Value = 0
    
    If rst.RecordCount <> 0 Then
        lblTotReg.Caption = rst.RecordCount
        
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
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = NulosC(rst("tc"))
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosC(rst("glosa"))
            
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(NulosC(rst("totdeb")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Format(NulosC(rst("tothab")), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Format(NulosN(rst("totsal")), FORMAT_MONTO)
            
            
            Fg1.TextMatrix(Fg1.Rows - 1, 14) = NulosN(rst("idmes"))
            
            If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 11)) < 0 Then '--
                Fg1.TextMatrix(Fg1.Rows - 1, 12) = "Si"
                Fg1.TextMatrix(Fg1.Rows - 1, 13) = ""
                sPer = sPer + Abs(NulosN(rst("totsal")))
            Else
                Fg1.TextMatrix(Fg1.Rows - 1, 12) = ""
                Fg1.TextMatrix(Fg1.Rows - 1, 13) = "Si"
                sGan = sGan + NulosN(rst("totsal"))
            End If
            
            lblGan.Caption = Format(sGan, FORMAT_MONTO)
            lblPer.Caption = Format(sPer, FORMAT_MONTO)
            
            '---avaluar si el saldo es =0 then
            
            If NulosN(Format(Fg1.TextMatrix(Fg1.Rows - 1, 11), "0.000")) = 0 Then
                Fg1.Rows = Fg1.Rows - 1
            End If
            
            
Seguir:
            
            rst.MoveNext
        Loop
        
    
    End If
    Frame1.Visible = False
    Set rst = Nothing
    Me.MousePointer = vbDefault


End Sub

Function Grabar() As Boolean
'Modificado 28/06/11 Johan Castro
'Agregar lineas de código para escribir el asientos de transferencia cuando se trate de una pérdida por redondeo a centimos

    Dim A, B, Rpta As Integer
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
    Dim TipCam As Double
    
    Dim IdCtaDestDeb As Long
    Dim IdCtaDestHab As Long
    
    Dim sSaldo As Double
    
    
    Dim rst As New ADODB.Recordset
    
    On Error GoTo LaCague
    
    RST_Busq rst, "SELECT mae_redondeo.* FROM mae_redondeo WHERE mae_redondeo.idmon=" & NulosC(LblIdMon.Caption) & " and mae_redondeo.idlib = " & NulosN(LblIdLibro.Caption) & " ;", xCon
    IdCtaGanancia = NulosN(rst("idcuengan"))
    IdCtaPerdida = NulosN(rst("idcuenper"))
    Set rst = Nothing
    
    RST_Busq rst, "SELECT con_planctas.ctadesdeb, con_planctas.ctadeshab FROM con_planctas WHERE (((con_planctas.id)=" & IdCtaPerdida & "));", xCon
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
    
    
    Frame1.Visible = True
    ProgressBar2.Max = Fg1.Rows - 1
    ProgressBar2.Min = 0
    
    For mRow = 2 To Fg1.Rows - 1
        ProgressBar2.Value = mRow
        DoEvents
        
        mIdMes = NulosN(Fg1.TextMatrix(mRow, 14))
        sSaldo = NulosN(Fg1.TextMatrix(mRow, 11))
        mIdMov = NulosN(Fg1.TextMatrix(mRow, 1))
        
        xNumAsiento = DevuelveNumAsiento(NulosN(LblIdLibro.Caption), mIdMov, mIdMes, xCon)
        TipCam = NulosN(Fg1.TextMatrix(mRow, 7))
        
        '--eliminamos los registros anteriormente grabados
        xCon.Execute "delete from con_diario where idlib = " & NulosN(LblIdLibro.Caption) & " and idmov = " & mIdMov & " and ridtipper=999"
        
        
        '**************************************************************************************
        If sSaldo = 0 Then
            MsgBox "SSS"
        End If
        
        RstDia.AddNew
        RstDia("año") = AnoTra
        RstDia("idmes") = mIdMes
        RstDia("idlib") = NulosN(LblIdLibro.Caption)
        RstDia("idmov") = mIdMov
        RstDia("numasi") = xNumAsiento
        RstDia("tc") = TipCam
        
        RstDia("impdebsol") = 0
        RstDia("impdebdol") = 0
        RstDia("imphabsol") = 0
        RstDia("imphabdol") = 0
                
        If NulosN(LblIdMon.Caption) = 2 Then
            '--si es perdida
            If LCase(Fg1.TextMatrix(mRow, 12)) = "si" Then
                RstDia("idcue") = IdCtaPerdida
                RstDia("impdebdol") = Abs(sSaldo)
            Else
                '--si es ganancia
                RstDia("idcue") = IdCtaGanancia
                RstDia("imphabdol") = Abs(sSaldo)
            End If
        
        Else
            If LCase(Fg1.TextMatrix(mRow, 12)) = "si" Then
                RstDia("idcue") = IdCtaPerdida
                RstDia("impdebsol") = Abs(sSaldo)
            Else
                RstDia("idcue") = IdCtaGanancia
                RstDia("imphabsol") = Abs(sSaldo)
            End If
        End If
        
        RstDia("fchasi") = CDate("01/" + Format(mIdMes, "00") + "/" + AnoTra)
        RstDia("fchdoc") = CDate(Fg1.TextMatrix(mRow, 5))
        
        RstDia("idmon") = NulosN(LblIdMon.Caption)
        
        RstDia("ridlib") = NulosN(LblIdLibro.Caption)
        RstDia("iddocpro") = 0
        RstDia("ridtipper") = 999
        RstDia("ridper") = 0
        RstDia("rtipdoc") = 0
        RstDia("rfchope") = Null
        RstDia("rnumerodoc") = Null
        RstDia("rregistro") = Fg1.TextMatrix(mRow, 3)
        RstDia("rglosaope") = NulosC(Fg1.TextMatrix(mRow, 8))
        RstDia("ridmon") = NulosN(LblIdMon.Caption)
        RstDia("ajuste") = 0
        RstDia("aplicatc") = -1
        RstDia.Update
        
        
        '**************************************************************************************
        
            '--destinos de perdida
           
            If IdCtaDestDeb <> 0 And IdCtaDestHab <> 0 And Fg1.TextMatrix(mRow, 12) = "Si" Then
                '************************************************************************************************
                '--destinos automatico cta debe
                RstDia.AddNew
                RstDia("año") = AnoTra
                RstDia("idmes") = mIdMes
                RstDia("idlib") = NulosN(LblIdLibro.Caption)
                RstDia("idmov") = mIdMov
                RstDia("numasi") = xNumAsiento
                RstDia("tc") = TipCam
                RstDia("iddoc") = 0
                RstDia("impdebsol") = 0
                RstDia("impdebdol") = 0
                RstDia("imphabsol") = 0
                RstDia("imphabdol") = 0
                RstDia("idcue") = IdCtaDestDeb
                If NulosN(LblIdMon.Caption) = 2 Then
                    RstDia("impdebdol") = Abs(sSaldo)
                Else
                    RstDia("impdebsol") = Abs(sSaldo)
                End If
                
                RstDia("fchasi") = CDate("01/" + Format(mIdMes, "00") + "/" + AnoTra)
                RstDia("fchdoc") = CDate(Fg1.TextMatrix(mRow, 5))
                
                RstDia("idmon") = NulosN(LblIdMon.Caption)
                
                RstDia("ridlib") = NulosN(LblIdLibro.Caption)
                RstDia("iddocpro") = 0
                RstDia("ridtipper") = 999
                RstDia("ridper") = 0
                RstDia("rtipdoc") = 0
                RstDia("rfchope") = Null
                RstDia("rnumerodoc") = Null
                RstDia("rregistro") = Fg1.TextMatrix(mRow, 3)
                RstDia("rglosaope") = NulosC(Fg1.TextMatrix(mRow, 8))
                RstDia("ridmon") = NulosN(LblIdMon.Caption)
                RstDia("ajuste") = 0
                RstDia("aplicatc") = -1
                
                RstDia.Update
                
                '************************************************************************************************
                '--destinos automatico cta debe
                RstDia.AddNew
                RstDia("año") = AnoTra
                RstDia("idmes") = mIdMes
                RstDia("idlib") = NulosN(LblIdLibro.Caption)
                RstDia("idmov") = mIdMov
                RstDia("numasi") = xNumAsiento
                RstDia("tc") = TipCam
                RstDia("iddoc") = 0
                RstDia("impdebsol") = 0
                RstDia("impdebdol") = 0
                RstDia("imphabsol") = 0
                RstDia("imphabdol") = 0
                RstDia("idcue") = IdCtaDestHab
                
                If NulosN(LblIdMon.Caption) = 2 Then
                    RstDia("imphabdol") = Abs(sSaldo)
                Else
                    RstDia("imphabsol") = Abs(sSaldo)
                End If
                
                RstDia("fchasi") = CDate("01/" + Format(mIdMes, "00") + "/" + AnoTra)
                RstDia("fchdoc") = CDate(Fg1.TextMatrix(mRow, 5))
                
                RstDia("idmon") = NulosN(LblIdMon.Caption)
                
                RstDia("ridlib") = NulosN(LblIdLibro.Caption)
                RstDia("iddocpro") = 0
                RstDia("ridtipper") = 999
                RstDia("ridper") = 0
                RstDia("rtipdoc") = 0
                RstDia("rfchope") = Null
                RstDia("rnumerodoc") = Null
                RstDia("rregistro") = Fg1.TextMatrix(mRow, 3)
                RstDia("rglosaope") = NulosC(Fg1.TextMatrix(mRow, 8))
                RstDia("ridmon") = NulosN(LblIdMon.Caption)
                RstDia("ajuste") = 0
                RstDia("aplicatc") = -1

                RstDia.Update
                '************************************************************************************************
                
            End If
            '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        
        
        
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




Private Sub VerLineal()
    Dim xRstLineal As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLDiario As String
    Dim xBase As String
    Dim nSQLPer As String
    Dim nSQLFecha As String
    Dim nSQLDiario1 As String
    Dim xImpSaldo As Double
    Dim xIdLibro As Integer
    Dim nSQLWhere As String
    Dim nSQLApertura As String '--filtro para documentos de apertura
    Dim xEliminaReg As Boolean
    Dim nSQLDocNCBancos As String '--almacenara los documentos de NC que pasan x banco
   '--2 Ventas
   '--1 Compras
   '--9 Boleta
   
   '-----------------------------------------------
    
    xIdLibro = LblIdLibro.Caption
    Select Case xIdLibro
        Case 1, 4, 40, 9
            xIdLibroRef = "6,8,39"
        Case 2 '--ventas
            xIdLibroRef = "5,6,8,37"
        Case 37 '--Letras
            xIdLibroRef = "6,42"
        Case 41 '--Lgd
            xIdLibroRef = "6,41"
        Case 42 '--Planilla Letras
            xIdLibroRef = "6"
        Case 999
'''            If OptReem1.Value = True Then
                xIdLibroRef = "6"
'''            Else
'''                xIdLibroRef = "41"
'''            End If
    End Select
    '-----------------------------------------------
    Fg1.Rows = Fg1.FixedRows
    DoEvents
   
    '--------------------------------------------
    '--con_diario.tipmov(1=Ingresos; 2=Egresos)
    '--con_diario.tipo  (1=Origen;   2=Destino)
    '--con_diario.rtipdoc=7(Nota de Credito)
    If xIdLibro = 1 Or xIdLibro = 4 Or xIdLibro = 9 Or xIdLibro = 40 Or xIdLibro = 999 Then
    '--compras, honorarios, reembolsables, boleta pago,percepciones
        xBase = vbCr & " IIf(con_diario.tipmov =1, IIf(con_diario.tipo =1,IIf(con_diario.rtipdoc=7,-1,1),IIf(con_diario.rtipdoc=7,1,-1)) , IIf(con_diario.tipo in (1),IIf(con_diario.rtipdoc=7,1,-1),1) ) as xbase, "
    Else
    '--ventas, lgd, letras, planilla letras
        xBase = " IIf(con_diario.tipmov =1, IIf(con_diario.tipo =1,IIf(con_diario.rtipdoc=7,1,-1),1), IIf(con_diario.tipo in (0,1),1,IIf(con_diario.rtipdoc=7,1,-1)) ) as xbase, "
    End If
    '--------------------------------------------
''    If OptFch(0).Value = True Then '--x fecha de documento
''        nSQLFecha = " and ( vta_ventas.fchdoc between CDate('" & TxtFchIni.Valor & "') and CDate('" & TxtFchFin.Valor & "') )"
''    ElseIf OptFch(1).Value = True Then '--x fecha de registro
''        nSQLFecha = " and ( vta_ventas.fchreg between CDate('" & TxtFchIni.Valor & "') and CDate('" & TxtFchFin.Valor & "') )"
''    End If
    
        
    nSQLPer = " and con_diario.ridper =" & NulosN(LblIdCliPro.Caption)
    
    nSQLIdMon = " vta_ventas.idmon=" & LblIdMon.Caption
    '--------------------------------------------
    nSQLWhere = nSQLFecha & nSQLPer & nSQLApertura & nSQLIdMon
    '--------------------------------------------
    
    
    If xIdLibro = 1 Then
        
        '--Verificar si hay documentos de NC que fueron registrados en Tesoreria Ingresos - Egresos
        nSQLDocNCBancos = BuscarNCBancos()
        If nSQLDocNCBancos <> "" Then
            nSQLDocNCBancos = " and com_compras.id not in (" & nSQLDocNCBancos & ")"
        End If
        
        '--Cancelacion de compras con nota de credito excepto nc que se registran en tesoreria
        nSQLDiario1 = " UNION " _
            + vbCr + " SELECT Left(com_compras_1.numreg,2) & mae_libros_1.codsun & Right(com_compras_1.numreg,4) AS rregistro, Mid(com_compras!numreg,1,2) & mae_libros!codsun & Mid(com_compras.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_prov.nombre AS razonsocial, mae_documento.abrev, " _
            + vbCr + " com_compras.numser & '-' & com_compras.numdoc AS numdoc, com_compras.fchdoc AS fchemi, mae_moneda.simbolo, " _
            + vbCr + " IIf(com_compras_1.tc=0,con_tc.impven,com_compras_1.tc) AS tipcam, 2 AS tipmov, 1 AS tipo, 1 AS xbase, com_compras.imptot AS imptotal, " _
            + vbCr + " IIf(com_compras.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, " _
            + vbCr + " IIf(com_compras.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) AS imptotdol, " _
            + vbCr + " com_compras.idpro AS ridper, com_compras_1.numser & '-' & com_compras_1.numdoc AS numdoc2, com_compras.glosa AS glosaope, com_compras_1.id AS iddoc, com_compras.idmon " _
            + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((((com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) INNER JOIN com_compras AS com_compras_1 ON com_compras.iddocref = com_compras_1.id)  " _
            + vbCr + " LEFT JOIN mae_libros AS mae_libros_1 ON com_compras_1.idlib = mae_libros_1.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) LEFT JOIN mae_prov ON com_compras.idpro = mae_prov.id " _
            + vbCr + " WHERE (com_compras.iddocref<>0 ) " & Replace(nSQLPer, "con_diario.ridper", "com_compras.idpro") & nSQLDocNCBancos

'''        '--Cancelacion de notas de credito con compras excepto nc que se registran en tesoreria
'''        nSQLDocNCBancos = Replace(nSQLDocNCBancos, "com_compras", "com_compras_1")
'''
'''        nSQLDiario1 = nSQLDiario1 _
'''            + vbCr + " UNION " _
'''            + vbCr + " SELECT Left(com_compras_1.numreg,2) & mae_libros_1.codsun & Right(com_compras_1.numreg,4) AS rregistro, Mid(com_compras!numreg,1,2) & mae_libros!codsun & Mid(com_compras.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_prov.nombre AS razonsocial, mae_documento.abrev, " _
'''            + vbCr + " com_compras.numser & '-' & com_compras.numdoc AS numdoc, com_compras.fchdoc AS fchemi, mae_moneda.simbolo, " _
'''            + vbCr + " IIf(com_compras_1.tc=0,con_tc.impven,com_compras_1.tc) AS tipcam, 2 AS tipmov, 1 AS tipo, 1 AS xbase, com_compras_1.imptot AS imptotal, " _
'''            + vbCr + " IIf(com_compras.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, " _
'''            + vbCr + " IIf(com_compras.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) AS imptotdol, " _
'''            + vbCr + " com_compras.idpro AS ridper, com_compras_1.numser & '-' & com_compras_1.numdoc AS numdoc2, com_compras.glosa AS glosaope, com_compras_1.id AS iddoc, com_compras.idmon " _
'''            + vbCr + " FROM (com_compras AS com_compras_1 LEFT JOIN mae_libros AS mae_libros_1 ON com_compras_1.idlib = mae_libros_1.id) INNER JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (((com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) " _
'''            + vbCr + " LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha) LEFT JOIN mae_prov ON com_compras.idpro = mae_prov.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON com_compras_1.iddocref = com_compras.id " _
'''            + vbCr + " WHERE (com_compras_1.iddocref<>0 ) " & Replace(nSQLPer, "con_diario.ridper", "com_compras.idpro") & nSQLDocNCBancos
        '--------------------------------------------------
    
    ElseIf xIdLibro = 2 Then
        '--Verificar si hay documentos de NC que fueron registrados en Tesoreria Ingresos - Egresos
        nSQLDocNCBancos = BuscarNCBancos()
        If nSQLDocNCBancos <> "" Then
            nSQLDocNCBancos = " and vta_ventas.id not in (" & nSQLDocNCBancos & ")"
        End If
        
        '--Cancelacion de ventas con nota de credito excepto nc que se registran en tesoreria
        nSQL = nSQL _
            + vbCr + " UNION " _
            + vbCr + " SELECT Left(vta_ventas_1.numreg,2) & mae_libros_1.codsun & Right(vta_ventas_1.numreg,4) AS rregistro, Mid(vta_ventas!numreg,1,2) & mae_libros!codsun & Mid(vta_ventas.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_prov.nombre AS razonsocial, mae_documento.abrev, " _
            + vbCr + " vta_ventas.numser & '-' & vta_ventas.numdoc AS numdoc, vta_ventas.fchdoc AS fchemi, mae_moneda.simbolo, IIf(vta_ventas_1.tc=0,con_tc.impven,vta_ventas_1.tc) AS tipcam, " _
            + vbCr + " 2 AS tipmov, 1 AS tipo, 1 AS xbase, vta_ventas.imptotdoc AS imptotal, " _
            + vbCr + " IIf(vta_ventas.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, " _
            + vbCr + " IIf(vta_ventas.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) AS imptotdol, " _
            + vbCr + " vta_ventas.idcli AS ridper, vta_ventas_1.numser & '-' & vta_ventas_1.numdoc AS numdoc2, vta_ventas.glosa AS glosaope, vta_ventas_1.id AS iddoc, vta_ventas.idmon " _
            + vbCr + " FROM (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((((vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) INNER JOIN vta_ventas AS vta_ventas_1 ON vta_ventas.iddocref = vta_ventas_1.id)  " _
            + vbCr + " LEFT JOIN mae_libros AS mae_libros_1 ON vta_ventas_1.idlib = mae_libros_1.id) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) LEFT JOIN mae_prov ON vta_ventas.idcli = mae_prov.id " _
            + vbCr + " WHERE (vta_ventas.iddocref<>0 ) " & Replace(nSQLPer, "con_diario.ridper", "vta_ventas.idcli") & nSQLDocNCBancos

'''        '--Cancelacion de notas de credito con ventas excepto nc que se registran en tesoreria
'''        nSQLDocNCBancos = Replace(nSQLDocNCBancos, "vta_ventas", "vta_ventas_1")
'''
'''        nSQL = nSQL _
'''            + vbCr + " UNION " _
'''            + vbCr + " SELECT Left(vta_ventas_1.numreg,2) & mae_libros_1.codsun & Right(vta_ventas_1.numreg,4) AS rregistro, Mid(vta_ventas!numreg,1,2) & mae_libros!codsun & Mid(vta_ventas.numreg,3,4) AS registro, mae_libros.descripcion AS libro, mae_prov.nombre AS razonsocial, mae_documento.abrev, " _
'''            + vbCr + " vta_ventas.numser & '-' & vta_ventas.numdoc AS numdoc, vta_ventas.fchdoc AS fchemi, mae_moneda.simbolo, IIf(vta_ventas_1.tc=0,con_tc.impven,vta_ventas_1.tc) AS tipcam,  " _
'''            + vbCr + " 2 AS tipmov, 1 AS tipo, 1 AS xbase, vta_ventas_1.imptotdoc AS imptotal, " _
'''            + vbCr + " IIf(vta_ventas.idmon=1,imptotal,imptotal*tipcam) AS imptotsol, " _
'''            + vbCr + " IIf(vta_ventas.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) AS imptotdol, " _
'''            + vbCr + " vta_ventas.idcli AS ridper, vta_ventas_1.numser & '-' & vta_ventas_1.numdoc AS numdoc2, vta_ventas.glosa AS glosaope, vta_ventas_1.id AS iddoc, vta_ventas.idmon " _
'''            + vbCr + " FROM (vta_ventas AS vta_ventas_1 LEFT JOIN mae_libros AS mae_libros_1 ON vta_ventas_1.idlib = mae_libros_1.id) INNER JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (((vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) " _
'''            + vbCr + " LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_prov ON vta_ventas.idcli = mae_prov.id) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON vta_ventas_1.iddocref = vta_ventas.id " _
'''            + vbCr + " WHERE (vta_ventas_1.iddocref<>0 ) " & Replace(nSQLPer, "con_diario.ridper", "vta_ventas.idcli") & nSQLDocNCBancos
        '--------------------------------------------------
    
    End If
    
    nSQLDiario = " SELECT Last(xdet.rregistro) as xrregistro, xdet.iddoc, last(xdet.simbolo) as xsimbolo, Sum(xdet.imptotsol) AS xtotsol, Sum(xdet.imptotdol) AS xtotdol, Last(xdet.registro) AS xregistro, Last(xdet.fchemi) AS xfchcancel, Count(xdet.tipmov) AS xcanreg,last(xdet.rglosaope) as xglosa ,last(xdet.razonsocial) as xnombre, last(xdet.numdoc) as xnumdoc,last(xdet.abrev) as xabrev  " _
        + vbCr + " FROM ( " _
        + vbCr + " SELECT con_diario.rregistro, Format([con_diario].[idmes],'00') & [mae_libros].[codsun] & Format([con_diario].[numasi],'0000') AS registro, mae_libros.descripcion AS libro, IIf([con_diario].[ridtipper2]=5,[mae_bancos].[abrev],IIf([con_diario].[ridtipper2]=2,[mae_cliente].[nombre],IIf([con_diario].[ridtipper2]=1,[mae_prov].[nombre],''))) AS razonsocial, tes_documentos.abrev, con_diario.rnumerodoc2 AS numdoc, con_diario.rfchope2 AS fchemi, mae_moneda.simbolo, IIf([con_diario].[aplicatc]=0,[con_tc].[impven],[con_diario].[tc]) AS tipcam, " _
        + vbCr + " con_diario.tipmov, con_diario.tipo, " & xBase _
        + vbCr + " IIf(con_diario.idmon=1,(con_diario.impdebsol+con_diario.imphabsol),(con_diario.impdebdol+con_diario.imphabdol)) AS imptotal, " _
        + vbCr + " IIf(con_diario.idmon=1,imptotal,imptotal*tipcam) * xbase  AS imptotsol, " _
        + vbCr + " IIf(con_diario.idmon=2,imptotal,IIf(tipcam=0,0,imptotal/tipcam)) * xbase AS imptotdol, " _
        + vbCr + " con_diario.ridper, con_diario.rnumerodoc AS numdoc2, con_diario.rglosaope, con_diario.iddoc, con_diario.idmon  " _
        + vbCr + " FROM ((((((con_diario LEFT JOIN mae_moneda ON con_diario.idmon = mae_moneda.id) LEFT JOIN con_tc ON con_diario.fchdoc = con_tc.fecha) LEFT JOIN tes_documentos ON con_diario.rtipdoc2 = tes_documentos.id) LEFT JOIN mae_libros ON con_diario.idlib = mae_libros.id) LEFT JOIN mae_bancos ON con_diario.ridper2 = mae_bancos.id) LEFT JOIN mae_cliente ON con_diario.ridper2 = mae_cliente.id) LEFT JOIN mae_prov ON con_diario.ridper2 = mae_prov.id " _
        + vbCr + " WHERE (((con_diario.idlib) In (" & xIdLibroRef & ")) AND ((con_diario.ridlib)=" & xIdLibro & ")) " & nSQLPer _
        + vbCr + nSQLDiario1 _
        + vbCr + " ) AS xdet " _
        + vbCr + " GROUP BY xdet.iddoc "

    Select Case xIdLibro
        Case 1 '--compras
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "com_compras")
            nSQLWhere = Replace(nSQLWhere, "con_diario.ridper", "com_compras.idpro")
                        
            nSQL = "SELECT  com_compras.id as iddoc,com_compras.tipdoc, mae_prov.numruc, mae_prov.nombre AS nombre, IIf(com_compras.numreg Is Null Or com_compras.numreg='',mae_libros.codsun,Left(com_compras.numreg,2) & mae_libros.codsun & Right(com_compras.numreg,4)) AS registro, mae_documento.abrev, " _
                + vbCr + " com_compras.numser+'-'+com_compras.numdoc AS numdoc2, com_compras.fchdoc, mae_moneda.simbolo, " _
                + vbCr + " IIf(com_compras.tc Is Null Or com_compras.tc=0,con_tc.impven,com_compras.tc) AS tipcam, com_compras.idmon, IIf(com_compras.numreg='000001',com_compras.imptotori,com_compras.imptot) AS imptotal, com_compras.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(com_compras.idmon=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(com_compras.idmon=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, " _
                + vbCr + " com_compras.glosa , com_compras.numerodocref as docref, " _
                + vbCr + " xtotsol, xtotdol, IIf(com_compras.idmon=1,imptotal-iif(xtotsol is null,0,xtotsol),imptotal-iif(xtotdol is null,0,xtotdol)) AS saldo, xregistro, xnombre,xabrev,xnumdoc,xfchcancel,xglosa ,xcanreg " _
                + vbCr + " FROM ( (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN (mae_documento RIGHT JOIN (com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_documento.id = com_compras.tipdoc) ON mae_prov.id = com_compras.idpro) ON mae_moneda.id = com_compras.idmon) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
                + vbCr + " ) LEFT JOIN ( " & nSQLDiario & "" _
                + vbCr + " ) AS xpag ON com_compras.id = xpag.iddoc " _
                + vbCr + " WHERE (((IIf(com_compras.numreg='000001',com_compras.imptotori,com_compras.imptot))<>0)) " & nSQLWhere _
                + vbCr + " ORDER BY IIf(com_compras.numreg Is Null Or com_compras.numreg='',mae_libros.codsun,Left(com_compras.numreg,2) & mae_libros.codsun & Right(com_compras.numreg,4)), com_compras.fchdoc;"
                
        Case 999 '--Reembolsables
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "com_reembolsables")
            nSQLWhere = Replace(nSQLWhere, "con_diario.ridper", "com_reembolsables.idpro")
            nSQLWhere = Replace(nSQLWhere, "fchreg", "fchdoc")
            
            nSQL = "SELECT  com_reembolsables.id as iddoc,com_reembolsables.tipdoc, mae_prov.numruc, mae_prov.nombre AS nombre, IIf(com_reembolsables.numreg Is Null Or com_reembolsables.numreg='',mae_libros.codsun,Left(com_reembolsables.numreg,2) & mae_libros.codsun & Right(com_reembolsables.numreg,4)) AS registro, mae_documento.abrev, " _
                + vbCr + " com_reembolsables.numser+'-'+com_reembolsables.numdoc AS numdoc2, com_reembolsables.fchdoc, mae_moneda.simbolo, " _
                + vbCr + " IIf(com_reembolsables.tc Is Null Or com_reembolsables.tc=0,con_tc.impven,com_reembolsables.tc) AS tipcam, com_reembolsables.idmon, IIf(com_reembolsables.numreg='000001',com_reembolsables.imptotori,com_reembolsables.imptot) AS imptotal, com_reembolsables.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(com_reembolsables.idmon=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(com_reembolsables.idmon=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, " _
                + vbCr + " com_reembolsables.glosa , com_reembolsables.numerodocref as docref, " _
                + vbCr + " xtotsol, xtotdol, IIf(com_reembolsables.idmon=1,imptotal-xtotsol,imptotal-xtotdol) AS saldo, xregistro, xnombre,xabrev,xnumdoc,xfchcancel,xglosa ,xcanreg " _
                + vbCr + " FROM ( (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN (mae_documento RIGHT JOIN (com_reembolsables LEFT JOIN mae_libros ON com_reembolsables.idlib = mae_libros.id) ON mae_documento.id = com_reembolsables.tipdoc) ON mae_prov.id = com_reembolsables.idpro) ON mae_moneda.id = com_reembolsables.idmon) LEFT JOIN con_tc ON com_reembolsables.fchdoc = con_tc.fecha " _
                + vbCr + " ) LEFT JOIN ( " & nSQLDiario & "" _
                + vbCr + " ) AS xpag ON com_reembolsables.id = xpag.iddoc " _
                + vbCr + " WHERE (((com_reembolsables.tipdoc)<>7) AND ((IIf(com_reembolsables.numreg='000001',com_reembolsables.imptotori,com_reembolsables.imptot))<>0)) " & nSQLWhere _
                + vbCr + " ORDER BY IIf(com_reembolsables.numreg Is Null Or com_reembolsables.numreg='',mae_libros.codsun,Left(com_reembolsables.numreg,2) & mae_libros.codsun & Right(com_reembolsables.numreg,4)), com_reembolsables.fchdoc;"
        
        Case 2 '--ventas
            nSQLWhere = Replace(nSQLWhere, "con_diario.ridper", "vta_ventas.idcli")
            
            nSQL = "SELECT  vta_ventas.id as iddoc,vta_ventas.tipdoc, mae_cliente.numruc, mae_cliente.nombre AS nombre, IIf(vta_ventas.numreg Is Null Or vta_ventas.numreg='',mae_libros.codsun,Left(vta_ventas.numreg,2) & mae_libros.codsun & Right(vta_ventas.numreg,4)) AS registro, mae_documento.abrev, " _
                + vbCr + " vta_ventas.numser+'-'+vta_ventas.numdoc AS numdoc2, vta_ventas.fchdoc, mae_moneda.simbolo, " _
                + vbCr + " IIf(vta_ventas.tc Is Null Or vta_ventas.tc=0,con_tc.impven,vta_ventas.tc) AS tipcam, vta_ventas.idmon, IIf(vta_ventas.numreg='000001',vta_ventas.imptotori,vta_ventas.imptotdoc) AS imptotal, vta_ventas.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(vta_ventas.idmon=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(vta_ventas.idmon=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, " _
                + vbCr + " vta_ventas.glosa ,vta_ventas.numerodocref as docref,  " _
                + vbCr + " xtotsol, xtotdol, IIf(vta_ventas.idmon=1,imptotal-xtotsol,imptotal-xtotdol) AS saldo,xregistro, xnombre,xabrev,xnumdoc,xfchcancel,xglosa ,xcanreg " _
                + vbCr + " FROM ( (mae_moneda RIGHT JOIN (mae_cliente RIGHT JOIN (mae_documento RIGHT JOIN (vta_ventas LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) ON mae_documento.id = vta_ventas.tipdoc) ON mae_cliente.id = vta_ventas.idcli) ON mae_moneda.id = vta_ventas.idmon) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha " _
                + vbCr + " ) LEFT JOIN ( " & nSQLDiario & "" _
                + vbCr + " ) AS xpag ON vta_ventas.id = xpag.iddoc " _
                + vbCr + " WHERE vta_ventas.anulado=0 and (IIf(vta_ventas.numreg='000001',vta_ventas.imptotori,vta_ventas.imptotdoc)<>0) " & nSQLWhere _
                + vbCr + " ORDER BY IIf(vta_ventas.numreg Is Null Or vta_ventas.numreg='',mae_libros.codsun,Left(vta_ventas.numreg,2) & mae_libros.codsun & Right(vta_ventas.numreg,4)), vta_ventas.fchdoc;"
        
        Case 4 '--Percepciones
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "con_percepcion")
            nSQLWhere = Replace(nSQLWhere, "con_diario.ridper", "con_percepcion.idcli")
            
            nSQL = "SELECT con_percepcion.id AS iddoc, con_percepcion.tipdoc, mae_prov.numruc, mae_prov.nombre AS nombre, IIf([con_percepcion].[numreg] Is Null Or [con_percepcion].[numreg]='',[mae_libros].[codsun],Left([con_percepcion].[numreg],2) & [mae_libros].[codsun] & Right([con_percepcion].[numreg],4)) AS registro, mae_documento.abrev, " _
                + vbCr + " [con_percepcion].[numser]+'-'+[con_percepcion].[numdoc] AS numdoc2, con_percepcion.fchdoc, mae_moneda.simbolo, " _
                + vbCr + " IIf([con_percepcion].[tc] Is Null Or [con_percepcion].[tc]=0,[con_tc].[impven],[con_percepcion].[tc]) AS tipcam, con_percepcion.idmon, con_percepcion.imptotper AS imptotal, con_percepcion.impsal, " _
                + vbCr + " IIf([imptotal]=0,0,IIf([con_percepcion].[idmon]=1,[imptotal],IIf([tipcam] Is Null,0,[imptotal]*[tipcam]))) AS imptotsol, " _
                + vbCr + " IIf([imptotal]=0,0,IIf([con_percepcion].[idmon]=2,[imptotal],IIf([tipcam] Is Null,0,[imptotal]/[tipcam]))) AS imptotdol, " _
                + vbCr + " con_percepcion.glosa, '' AS docref, " _
                + vbCr + " xtotsol, xtotdol, IIf(con_percepcion.idmon=1,imptotal-xtotsol,imptotal-xtotdol) AS saldo, xregistro, xnombre,xabrev,xnumdoc,xfchcancel,xglosa ,xcanreg " _
                + vbCr + " FROM ( ((((con_percepcion LEFT JOIN mae_documento ON con_percepcion.tipdoc = mae_documento.id) LEFT JOIN mae_libros ON con_percepcion.idlib = mae_libros.id) LEFT JOIN con_tc ON con_percepcion.fchdoc = con_tc.fecha) LEFT JOIN mae_prov ON con_percepcion.idcli = mae_prov.id) LEFT JOIN mae_moneda ON con_percepcion.idmon = mae_moneda.id " _
                + vbCr + " ) LEFT JOIN ( " & nSQLDiario & "" _
                + vbCr + " ) AS xpag ON con_percepcion.id = xpag.iddoc " _
                + vbCr + " Where (((con_percepcion.imptotper) <> 0)) " & nSQLWhere _
                + vbCr + " ORDER BY IIf([con_percepcion].[numreg] Is Null Or [con_percepcion].[numreg]='',[mae_libros].[codsun],Left([con_percepcion].[numreg],2) & [mae_libros].[codsun] & Right([con_percepcion].[numreg],4)), con_percepcion.fchdoc  "

        Case 9 '--Planilla Pago
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "pla_boleta")
            nSQLWhere = Replace(nSQLWhere, "con_diario.ridper", "pla_boleta.idemp")
                
            nSQL = "SELECT  pla_boleta.id as iddoc,pla_boleta.iddoc as tipdoc, pla_empleados.numdoc as  numruc, pla_empleados.nombre AS nombre, IIf(pla_boleta.numreg Is Null Or pla_boleta.numreg='',mae_libros.codsun,Left(pla_boleta.numreg,2) & mae_libros.codsun & Right(pla_boleta.numreg,4)) AS registro, mae_documento.abrev, " _
                + vbCr + " pla_boleta.numser+'-'+pla_boleta.numdoc AS numdoc2, pla_boleta.fchdoc, mae_moneda.simbolo, " _
                + vbCr + " con_tc.impven AS tipcam, pla_boleta.idmon, pla_boleta.imptot AS imptotal, pla_boleta.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(pla_boleta.idmon=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(pla_boleta.idmon=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, " _
                + vbCr + " pla_boleta.glosa ,'' as docref, " _
                + vbCr + " xtotsol, xtotdol, IIf(pla_boleta.idmon=1,imptotal-iif(xtotsol is null,0,xtotsol),imptotal-iif(xtotdol is null,0,xtotdol)) AS saldo, xregistro, xnombre,xabrev,xnumdoc,xfchcancel,xglosa ,xcanreg " _
                + vbCr + " FROM ( (mae_moneda RIGHT JOIN (pla_empleados RIGHT JOIN (mae_documento RIGHT JOIN (pla_boleta LEFT JOIN mae_libros ON pla_boleta.idlib = mae_libros.id) ON mae_documento.id = pla_boleta.iddoc) ON pla_empleados.id = pla_boleta.idemp) ON mae_moneda.id = pla_boleta.idmon) LEFT JOIN con_tc ON pla_boleta.fchdoc = con_tc.fecha " _
                + vbCr + " ) LEFT JOIN ( " & nSQLDiario & "" _
                + vbCr + " ) AS xpag ON pla_boleta.id = xpag.iddoc " _
                + vbCr + " WHERE (((pla_boleta.iddoc)<>7)) " & nSQLWhere _
                + vbCr + " ORDER BY IIf(pla_boleta.numreg Is Null Or pla_boleta.numreg='',mae_libros.codsun,Left(pla_boleta.numreg,2) & mae_libros.codsun & Right(pla_boleta.numreg,4)), pla_boleta.fchdoc;"
            
        Case 37 '--Letras
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "let_letra")
            nSQLWhere = Replace(nSQLWhere, "con_diario.ridper", "let_letra.idclipro")
            nSQLWhere = Replace(nSQLWhere, "fchdoc", "fchemi")

            nSQL = "SELECT let_letradet.corr AS iddoc, let_letra.tipdoc, mae_cliente.numruc, mae_cliente.nombre AS nombre, IIf(let_letra.numreg Is Null Or let_letra.numreg='',mae_libros.codsun,Left(let_letra.numreg,2) & mae_libros.codsun & Right(let_letra.numreg,4)) AS registro, mae_documento.abrev, " _
                + vbCr + " let_letra.ano & ' ' & let_letradet.numdoc & ' ' & let_letradet.numser AS numdoc2, let_letra.fchemi AS fchdoc, mae_moneda.simbolo, " _
                + vbCr + " IIf(let_letra.tc Is Null Or let_letra.tc=0,con_tc.impven,let_letra.tc) AS tipcam, let_letra.idmon, let_letradet.implet AS imptotal, let_letradet.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(let_letra.idmon=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(let_letra.idmon=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, " _
                + vbCr + " let_letra.glosa ,let_letra.numerodocref as docref, " _
                + vbCr + " xtotsol, xtotdol, IIf(let_letra.idmon=1,imptotal-xtotsol,imptotal-xtotdol) AS saldo,xregistro, xnombre,xabrev,xnumdoc,xfchcancel,xglosa ,xcanreg " _
                + vbCr + " FROM ( (((mae_moneda RIGHT JOIN (mae_libros RIGHT JOIN (mae_documento RIGHT JOIN let_letra ON mae_documento.id = let_letra.tipdoc) ON mae_libros.id = let_letra.idlib) ON mae_moneda.id = let_letra.idmon) INNER JOIN let_letradet ON let_letra.id = let_letradet.idlet) LEFT JOIN mae_cliente ON let_letra.idclipro = mae_cliente.id) LEFT JOIN con_tc ON let_letra.fchemi = con_tc.fecha " _
                + vbCr + " ) LEFT JOIN ( " & nSQLDiario & "" _
                + vbCr + " ) AS xpag ON let_letradet.corr = xpag.iddoc " _
                + vbCr + " WHERE (((IIf(let_letra.numreg='000001',let_letradet.imptotori,let_letradet.implet))<>0)) " & nSQLWhere _
                + vbCr + " ORDER BY IIf(let_letra.numreg Is Null Or let_letra.numreg='',mae_libros.codsun,Left(let_letra.numreg,2) & mae_libros.codsun & Right(let_letra.numreg,4)), let_letra.fchemi "
        
        Case 40 '--Honorarios
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "com_honorarios")
            nSQLWhere = Replace(nSQLWhere, "con_diario.ridper", "com_honorarios.idpro")
            
            nSQL = "SELECT  com_honorarios.id as iddoc,com_honorarios.tipdoc, mae_prov.numruc, mae_prov.nombre AS nombre, IIf(com_honorarios.numreg Is Null Or com_honorarios.numreg='',mae_libros.codsun,Left(com_honorarios.numreg,2) & mae_libros.codsun & Right(com_honorarios.numreg,4)) AS registro, mae_documento.abrev, " _
                + vbCr + " com_honorarios.numser+'-'+com_honorarios.numdoc AS numdoc2, com_honorarios.fchdoc, mae_moneda.simbolo, " _
                + vbCr + " IIf(com_honorarios.tc Is Null Or com_honorarios.tc=0,con_tc.impven,com_honorarios.tc) AS tipcam, com_honorarios.idmon, IIf(com_honorarios.numreg='000001',com_honorarios.imptotori,com_honorarios.imptot) AS imptotal, com_honorarios.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(com_honorarios.idmon=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(com_honorarios.idmon=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, " _
                + vbCr + " com_honorarios.glosa ,'' as docref, " _
                + vbCr + " xtotsol, xtotdol, IIf(com_honorarios.idmon=1,imptotal-xtotsol,imptotal-xtotdol) AS saldo, xregistro, xnombre,xabrev,xnumdoc,xfchcancel,xglosa ,xcanreg " _
                + vbCr + " FROM ( (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN (mae_documento RIGHT JOIN (com_honorarios LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) ON mae_documento.id = com_honorarios.tipdoc) ON mae_prov.id = com_honorarios.idpro) ON mae_moneda.id = com_honorarios.idmon) LEFT JOIN con_tc ON com_honorarios.fchdoc = con_tc.fecha " _
                + vbCr + " ) LEFT JOIN ( " & nSQLDiario & "" _
                + vbCr + " ) AS xpag ON com_honorarios.id = xpag.iddoc " _
                + vbCr + " WHERE (((com_honorarios.tipdoc)<>7) AND ((IIf(com_honorarios.numreg='000001',com_honorarios.imptotori,com_honorarios.imptot))<>0)) " & nSQLWhere _
                + vbCr + " ORDER BY IIf(com_honorarios.numreg Is Null Or com_honorarios.numreg='',mae_libros.codsun,Left(com_honorarios.numreg,2) & mae_libros.codsun & Right(com_honorarios.numreg,4)), com_honorarios.fchdoc"
    
    Case 41 '--Lgd
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "vta_gastodebito")
            nSQLWhere = Replace(nSQLWhere, "con_diario.ridper", "vta_gastodebito.idcli")
            
            nSQL = "SELECT  vta_gastodebito.id as iddoc,vta_gastodebito.tipdoc, mae_cliente.numruc, mae_cliente.nombre AS nombre, IIf(vta_gastodebito.numreg Is Null Or vta_gastodebito.numreg='',mae_libros.codsun,Left(vta_gastodebito.numreg,2) & mae_libros.codsun & Right(vta_gastodebito.numreg,4)) AS registro, mae_documento.abrev, " _
                + vbCr + " vta_gastodebito.numser+'-'+vta_gastodebito.numdoc AS numdoc2, vta_gastodebito.fchemi as fchdoc, mae_moneda.simbolo, " _
                + vbCr + " IIf(vta_gastodebito.tc Is Null Or vta_gastodebito.tc=0,con_tc.impven,vta_gastodebito.tc) AS tipcam, vta_gastodebito.idmon, vta_gastodebito.imptot AS imptotal, vta_gastodebito.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(vta_gastodebito.idmon=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(vta_gastodebito.idmon=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, " _
                + vbCr + " vta_gastodebito.glosa ,vta_gastodebito.numerodocref as docref,  " _
                + vbCr + " xtotsol, xtotdol, IIf(vta_gastodebito.idmon=1,imptotal-xtotsol,imptotal-xtotdol) AS saldo, xregistro, xnombre,xabrev,xnumdoc,xfchcancel,xglosa ,xcanreg " _
                + vbCr + " FROM ( (mae_moneda RIGHT JOIN (mae_cliente RIGHT JOIN (mae_documento RIGHT JOIN (vta_gastodebito LEFT JOIN mae_libros ON vta_gastodebito.idlib = mae_libros.id) ON mae_documento.id = vta_gastodebito.tipdoc) ON mae_cliente.id = vta_gastodebito.idcli) ON mae_moneda.id = vta_gastodebito.idmon) LEFT JOIN con_tc ON vta_gastodebito.fchemi = con_tc.fecha " _
                + vbCr + " ) LEFT JOIN ( " & nSQLDiario & "" _
                + vbCr + " ) AS xpag ON vta_gastodebito.id = xpag.iddoc " _
                + vbCr + " WHERE (((vta_gastodebito.tipdoc)<>7) AND ((IIf(vta_gastodebito.numreg='000001',vta_gastodebito.imptot,vta_gastodebito.imptot))<>0)) " & nSQLWhere _
                + vbCr + " ORDER BY IIf(vta_gastodebito.numreg Is Null Or vta_gastodebito.numreg='',mae_libros.codsun,Left(vta_gastodebito.numreg,2) & mae_libros.codsun & Right(vta_gastodebito.numreg,4)), vta_gastodebito.fchemi "
    
    
    Case 42 '--Planilla Letras
            nSQLWhere = Replace(nSQLWhere, "vta_ventas", "let_planilla")
            nSQLWhere = Replace(nSQLWhere, "con_diario.ridper", "mae_bancos.id")
            nSQLWhere = Replace(nSQLWhere, "fchdoc", "fchemi")
    
            nSQL = "SELECT let_planilla.id as iddoc, let_planilla.tipdoc, mae_bancos.numruc, mae_bancos.descripcion AS nombre, IIf(let_planilla.numreg Is Null Or let_planilla.numreg='',mae_libros.codsun,Left(let_planilla.numreg,2) & mae_libros.codsun & Right(let_planilla.numreg,4)) AS registro, mae_documento.abrev, " _
                + vbCr + " IIf(let_planilla.numser is null,'',let_planilla.numser & '-') & let_planilla.numdoc AS numdoc2, let_planilla.fchemi AS fchdoc, mae_moneda.simbolo, " _
                + vbCr + " IIf(let_planilla.tc Is Null Or let_planilla.tc=0,con_tc.impven,let_planilla.tc) AS tipcam, let_planilla.idmon, IIf(let_planilla.numreg='000001',let_planilla.imptot,let_planilla.imptot) AS imptotal, let_planilla.impsal, " _
                + vbCr + " IIf(imptotal=0,0,IIf(let_planilla.idmon=1,imptotal,IIf(tipcam Is Null,0,imptotal*tipcam))) AS imptotsol, " _
                + vbCr + " IIf(imptotal=0,0,IIf(let_planilla.idmon=2,imptotal,IIf(tipcam Is Null,0,imptotal/tipcam))) AS imptotdol, " _
                + vbCr + " let_planilla.glosa ,'' as docref, " _
                + vbCr + " xtotsol, xtotdol, IIf([let_planilla].[idmon]=1,[imptotal]-xtotsol,[imptotal]-xtotdol) AS saldo, xregistro, xnombre,xabrev,xnumdoc,xfchcancel,xglosa ,xcanreg " _
                + vbCr + " FROM ( (((mae_bancos RIGHT JOIN (mae_banconumcta RIGHT JOIN (mae_documento RIGHT JOIN let_planilla ON mae_documento.id = let_planilla.tipdoc) ON mae_banconumcta.id = let_planilla.idbcocta) ON mae_bancos.id = mae_banconumcta.idban) LEFT JOIN mae_libros ON let_planilla.idlib = mae_libros.id) LEFT JOIN con_tc ON let_planilla.fchemi = con_tc.fecha) LEFT JOIN mae_moneda ON let_planilla.idmon = mae_moneda.id " _
                + vbCr + " ) LEFT JOIN ( " & nSQLDiario & "" _
                + vbCr + " ) AS xpag ON let_planilla.id = xpag.iddoc " _
                + vbCr + " WHERE (((IIf(let_planilla.numreg='000001',let_planilla.imptot,let_planilla.imptot))<>0)) " & nSQLWhere _
                + vbCr + " ORDER BY IIf(let_planilla.numreg Is Null Or let_planilla.numreg='',mae_libros.codsun,Left(let_planilla.numreg,2) & mae_libros.codsun & Right(let_planilla.numreg,4)), let_planilla.fchemi "
    
    End Select

    Dim xFila As Long
    
    RST_Busq xRstLineal, nSQL, xCon
    
    xFila = Fg1.FixedRows
    If xRstLineal.State = 0 Then GoTo Salir
    If xRstLineal.RecordCount = 0 Then GoTo Salir
    
    '-----------------
    fraBarra.Visible = True
    fraBarra.Left = 2798
    fraBarra.Top = 2925
    
    ProgressBar1.Max = 1
    ProgressBar1.Value = 0
    
    fraBarra.Refresh
    ProgressBar1.Max = xRstLineal.RecordCount
    BAND_INTERRUMPIR = False
    '-----------------
    Fg2.Rows = Fg2.FixedRows
    Set xRstTot = Nothing
    PreparaRST
    DoEvents
    '-----------------
    Do While Not xRstLineal.EOF
        DoEvents
        
        If BAND_INTERRUMPIR = True Then GoTo Salir:
        ProgressBar1.Value = ProgressBar1.Value + 1
            
        Fg1.Rows = Fg1.Rows + 1
        
        xFila = Fg1.Rows - 1
                
        Fg1.TextMatrix(xFila, 1) = NulosC(xRstLineal("numruc"))
        Fg1.TextMatrix(xFila, 2) = NulosC(xRstLineal("nombre"))
        Fg1.TextMatrix(xFila, 3) = NulosC(xRstLineal("registro"))
        Fg1.TextMatrix(xFila, 4) = NulosC(xRstLineal("abrev"))
        Fg1.TextMatrix(xFila, 5) = NulosC(xRstLineal("numdoc2"))
        Fg1.TextMatrix(xFila, 6) = Format(NulosC(xRstLineal("fchdoc")), FORMAT_DATE)
        Fg1.TextMatrix(xFila, 7) = NulosC(xRstLineal("simbolo"))
        Fg1.TextMatrix(xFila, 8) = NulosC(xRstLineal("tipcam"))
        Fg1.TextMatrix(xFila, 9) = Format(NulosC(xRstLineal("imptotal")), FORMAT_MONTO)
        Fg1.TextMatrix(xFila, 10) = NulosC(xRstLineal("glosa"))
        
        Fg1.TextMatrix(xFila, 11) = NulosN(xRstLineal("xcanreg"))
        Fg1.TextMatrix(xFila, 12) = Format(NulosC(xRstLineal("xfchcancel")), FORMAT_DATE)
        Fg1.TextMatrix(xFila, 13) = NulosC(xRstLineal("xregistro"))
        Fg1.TextMatrix(xFila, 14) = NulosC(xRstLineal("xnombre"))
        Fg1.TextMatrix(xFila, 15) = NulosC(xRstLineal("xabrev"))
        Fg1.TextMatrix(xFila, 16) = NulosC(xRstLineal("xnumdoc"))
        
        If NulosN(xRstLineal("idmon")) = 1 Then
            Fg1.TextMatrix(xFila, 17) = Format(NulosN(xRstLineal("xtotsol")), FORMAT_MONTO)
        Else
            Fg1.TextMatrix(xFila, 17) = Format(NulosN(xRstLineal("xtotdol")), FORMAT_MONTO)
        End If
        Fg1.TextMatrix(xFila, 18) = NulosC(xRstLineal("xglosa"))
        
        If NulosN(xRstLineal("xcanreg")) <> 0 Then
            Fg1.TextMatrix(xFila, 19) = Format(NulosN(xRstLineal("saldo")), FORMAT_MONTO)
        Else
            Fg1.TextMatrix(xFila, 19) = Format(NulosC(xRstLineal("imptotal")), FORMAT_MONTO)
        End If
        xImpSaldo = NulosN(Fg1.TextMatrix(xFila, 19))
        
        Fg1.TextMatrix(xFila, 20) = NulosC(xRstLineal("iddoc"))
        Fg1.TextMatrix(xFila, 21) = NulosC(xRstLineal("idmon"))
        Fg1.TextMatrix(xFila, 22) = NulosC(xRstLineal("docref"))
        
        If xImpSaldo <> NulosN(xRstLineal("impsal")) Then
        
            '--Actualizar saldos a documento
            Select Case xIdLibro
                Case 1   '--Compras
                    xCon.Execute "Update com_compras set com_compras.impsal=" & xImpSaldo & " where com_compras.id = " & NulosC(xRstLineal("iddoc"))
                Case 4   '--Percepcion
                    xCon.Execute "Update con_percepcion set con_percepcion.impsal=" & xImpSaldo & " where con_percepcion.id = " & NulosC(xRstLineal("iddoc"))
                Case 40  '--Honorarios
                    xCon.Execute "Update com_honorarios set com_honorarios.impsal=" & xImpSaldo & " where com_honorarios.id = " & NulosC(xRstLineal("iddoc"))
                Case 2   '--Ventas
                    xCon.Execute "Update vta_ventas set vta_ventas.impsal=" & xImpSaldo & " where vta_ventas.id = " & NulosC(xRstLineal("iddoc"))
                Case 9  '--Boleta Pago
                    xCon.Execute "Update pla_boleta set pla_boleta.impsal=" & xImpSaldo & " where pla_boleta.id = " & NulosC(xRstLineal("iddoc"))
                Case 37  '--Letras
                    xCon.Execute "Update let_letradet set let_letradet.impsal=" & xImpSaldo & " where let_letradet.corr = " & NulosC(xRstLineal("iddoc"))
                Case 41  '--Lgd, Lgc
                    xCon.Execute "Update vta_gastodebito set vta_gastodebito.impsal=" & xImpSaldo & " where vta_gastodebito.id = " & NulosC(xRstLineal("iddoc"))
                Case 42  '--Planilla letras
                    xCon.Execute "Update let_planilla set let_planilla.impsal=" & xImpSaldo & " where let_planilla.id = " & NulosC(xRstLineal("iddoc"))
                Case 999 '--Reembolsables
                    xCon.Execute "Update com_reembolsables set com_reembolsables.impsal=" & xImpSaldo & " where com_reembolsables.id = " & NulosC(xRstLineal("iddoc"))
            End Select
        
        End If
        
        '--pintar las celdas
        If NulosN(xRstLineal("xcanreg")) = 0 Then
            '--Pendientes de agregar operaciones
            
        ElseIf NulosN(Format(xRstLineal("saldo"), FORMAT_MONTO)) = 0 Then
            '--Documentos Cancelados
            GRID_COLOR_FONDO Fg1, xFila, 1, xFila, Fg1.Cols - 1, &HA4FFA4
            
        ElseIf NulosN(Format(xRstLineal("saldo"), FORMAT_MONTO)) < 0 Then
            '--Documentos Observados
            GRID_COLOR_FONDO Fg1, xFila, 1, xFila, Fg1.Cols - 1, &H8C8CFF
            
        Else
            '--Documentos Pendientes
            GRID_COLOR_FONDO Fg1, xFila, 1, xFila, Fg1.Cols - 1, vbYellow '&H9BFFFF
        End If
        
        xEliminaReg = False
        
        If OptPen.Value = True Then
            If NulosN(xRstLineal("xcanreg")) <> 0 And NulosN(Format(xRstLineal("saldo"), FORMAT_MONTO)) = 0 Then
                Fg1.Rows = Fg1.Rows - 1
                xEliminaReg = True
            End If
        ElseIf OptCan.Value = True Then
            If NulosN(xRstLineal("xcanreg")) = 0 Or NulosN(Format(xRstLineal("saldo"), FORMAT_MONTO)) <> 0 Then
                Fg1.Rows = Fg1.Rows - 1
                xEliminaReg = True
            End If
        Else
            
        End If
        If xEliminaReg = False Then
            Acumular NulosC(xRstLineal("numruc")), NulosC(xRstLineal("nombre")), _
                     IIf(NulosN(xRstLineal("idmon")) = 1, NulosN(xRstLineal("imptotal")), 0), _
                     IIf(NulosN(xRstLineal("idmon")) = 2, NulosN(xRstLineal("imptotal")), 0), _
                     IIf(NulosN(xRstLineal("idmon")) = 1, NulosN(xRstLineal("xtotsol")), 0), _
                     IIf(NulosN(xRstLineal("idmon")) = 2, NulosN(xRstLineal("xtotdol")), 0)
        End If
        
        xRstLineal.MoveNext
    Loop
    
    '----------------
    '-- cargar datos del resumen
    xRstTot.Filter = ""
    If xRstTot.RecordCount <> 0 Then xRstTot.MoveFirst
    Do While Not xRstTot.EOF
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosC(xRstTot("ruc"))
        Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosC(xRstTot("nombre"))
        Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(NulosN(xRstTot("impmn")), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(NulosN(xRstTot("impme")), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(NulosN(xRstTot("pimpmn")), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 6) = Format(NulosN(xRstTot("pimpme")), FORMAT_MONTO)
        
        Fg2.TextMatrix(Fg2.Rows - 1, 7) = Format(NulosN(xRstTot("impmn")) - NulosN(xRstTot("pimpmn")), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 8) = Format(NulosN(xRstTot("impme")) - NulosN(xRstTot("pimpme")), FORMAT_MONTO)
        
        xRstTot.MoveNext
    Loop
    
    If Fg2.Rows > Fg2.FixedRows Then
        '--Ordenar
        GRID_ORDENAR Fg2, Fg2.FixedRows, 2
        
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(Fg2.Rows - 1, 2) = "TOTALES "
        Fg2.TextMatrix(Fg2.Rows - 1, 3) = Format(NulosN(GRID_SUMAR_COL(Fg2, 3)), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 4) = Format(NulosN(GRID_SUMAR_COL(Fg2, 4)), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 5) = Format(NulosN(GRID_SUMAR_COL(Fg2, 5)), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 6) = Format(NulosN(GRID_SUMAR_COL(Fg2, 6)), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 7) = Format(NulosN(GRID_SUMAR_COL(Fg2, 7)), FORMAT_MONTO)
        Fg2.TextMatrix(Fg2.Rows - 1, 8) = Format(NulosN(GRID_SUMAR_COL(Fg2, 8)), FORMAT_MONTO)

        With Fg2
            .Cell(flexcpForeColor, Fg2.Rows - 1, 1, Fg2.Rows - 1, Fg2.Cols - 1) = &H800000
            .Select Fg2.Rows - 1, 1, Fg2.Rows - 1, Fg2.Cols - 1
            .FillStyle = flexFillRepeat
            .CellFontBold = True
        End With
        
        '--Ajustar las columnas
        Fg2.AutoSizeMode = flexAutoSizeColWidth
        Fg2.AutoSize 3
        Fg2.AutoSize 4
        Fg2.AutoSize 5
        Fg2.AutoSize 6
        Fg2.AutoSize 7
        Fg2.AutoSize 8
        
    End If
    
Salir:

    Set xRstLineal = Nothing
    fraBarra.Visible = False
    BAND_INTERRUMPIR = False
End Sub


