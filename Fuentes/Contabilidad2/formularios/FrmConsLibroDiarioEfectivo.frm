VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmConsLibroDiarioEfectivo 
   Caption         =   "Consulta de Libro Caja y Bancos - Detalle de los Movimientos del Efectivo"
   ClientHeight    =   7320
   ClientLeft      =   825
   ClientTop       =   2925
   ClientWidth     =   11580
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10320
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsLibroDiarioEfectivo.frx":0000
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsLibroDiarioEfectivo.frx":0454
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsLibroDiarioEfectivo.frx":05C0
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsLibroDiarioEfectivo.frx":0B08
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Caption         =   "Moneda"
      Height          =   1125
      Left            =   4170
      TabIndex        =   15
      Top             =   450
      Width           =   1200
      Begin VB.OptionButton OptDol 
         Caption         =   "Dólares"
         Height          =   195
         Left            =   105
         TabIndex        =   3
         Top             =   720
         Width           =   930
      End
      Begin VB.OptionButton OptSol 
         Caption         =   "Soles"
         Height          =   195
         Left            =   105
         TabIndex        =   2
         Top             =   345
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   714
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cerrar"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame FraProgreso 
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   1140
      TabIndex        =   11
      Top             =   3855
      Visible         =   0   'False
      Width           =   6735
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   90
         TabIndex        =   12
         Top             =   390
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   -15
         X2              =   6735
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         X1              =   0
         X2              =   0
         Y1              =   -15
         Y2              =   1305
      End
      Begin VB.Label LblTituloProg 
         AutoSize        =   -1  'True
         Caption         =   "Procesando: Detalle de los Movimientos del Efectivo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   90
         Width           =   4380
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   2
         Left            =   5130
         TabIndex        =   13
         Top             =   90
         Width           =   1530
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cuenta(s)"
      Height          =   1125
      Left            =   5415
      TabIndex        =   10
      Top             =   450
      Width           =   6105
      Begin VSFlex7Ctl.VSFlexGrid Fg2 
         Height          =   855
         Left            =   90
         TabIndex        =   4
         Top             =   225
         Width           =   5940
         _cx             =   10477
         _cy             =   1508
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
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsLibroDiarioEfectivo.frx":0EA0
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
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   0
      TabIndex        =   7
      Top             =   450
      Width           =   4125
      Begin AspaTextBoxFecha.TextBoxFecha TxtFec2 
         Height          =   300
         Left            =   2730
         TabIndex        =   1
         Top             =   450
         Width           =   1275
         _ExtentX        =   2249
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
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFec1 
         Height          =   300
         Left            =   705
         TabIndex        =   0
         Top             =   450
         Width           =   1275
         _ExtentX        =   2249
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
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   450
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   2145
         TabIndex        =   8
         Top             =   450
         Width           =   420
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Align           =   2  'Align Bottom
      Height          =   5700
      Left            =   0
      TabIndex        =   5
      Top             =   1620
      Width           =   11580
      _cx             =   20426
      _cy             =   10054
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
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   2
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmConsLibroDiarioEfectivo.frx":0F04
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
   Begin VB.Menu mnumenu1 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mnu_insertcta 
         Caption         =   "Insertar Cuenta"
      End
      Begin VB.Menu mnulinea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuquitarcta 
         Caption         =   "Quitar Cuenta"
      End
   End
End
Attribute VB_Name = "FrmConsLibroDiarioEfectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vStrCons As String, vFormatString As String
Dim vStrFiltro1 As String, vFormatStrCuentas As String
Dim BAND_INTERRUMPIR As Boolean
'--VARIABLES PARA EXPORTAR A EXCEL
Dim Oleapp As Object

Private Sub ModifiTamanioColGridCta()
    If Fg2.Rows > 3 Then
        Fg2.ColWidth(2) = 3950
    ElseIf Fg2.Rows < 4 Then
        Fg2.ColWidth(2) = 4200
    End If
End Sub

Private Sub BuscarCta()
    Dim x As Long
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim vNext As Integer
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Num. Cuenta":  xCampos(0, 1) = "cuenta":        xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Desc Cuenta":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "6000":    xCampos(1, 3) = "C"
'    xCampos(2, 0) = "Codigo":       xCampos(2, 1) = "id":            xCampos(2, 2) = "1200":    xCampos(2, 3) = "N"
'    xCampos(3, 0) = "Tipo":         xCampos(3, 1) = "tipo":          xCampos(3, 2) = "1000":    xCampos(3, 3) = "N"
    
    xform.SQLCad = "SELECT cuenta, descripcion, id, tipo FROM con_planctas WHERE con_planctas.cuenta Like '10-[1,2]%'"
    
    xform.Titulo = "Buscando Cuenta Contable"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "cuenta"
    xform.CampoBusca = "cuenta"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
        
    If xRs.State = 1 Then
        On Error Resume Next
        If NulosN(xRs("tipo")) = 1 Then
            MsgBox "Seleccion incorreta...!", vbInformation, xTitulo
            Exit Sub
        End If
        'If Trim(xRs("cuenta")) = "10-1-01-01" Then
        If InStr(Trim(xRs("cuenta")), "10-1") > 0 Then
            vNext = 1
        Else
            If InStr(Trim(xRs("cuenta")), "10-2") > 0 Then
                vNext = 1
            Else
                vNext = 0
            End If
        End If
        If vNext = 0 Then
            MsgBox "Seleccion incorrecta...!", vbInformation, xTitulo
            Exit Sub
        End If
'        If Trim(NulosC(xRs("cuenta"))) = "10-1-01-01" Or Trim(NulosC(xRs("cuenta"))) = "10-2-01-01" Then
'    '        MsgBox "Seleccion incorrecta...!", vbInformation, xTitulo
'    '        Exit Sub
'        Else
'            MsgBox "Seleccion incorrecta...!", vbInformation, xTitulo
'            Exit Sub
'        End If
        
        If Trim(Fg2.TextMatrix(0, 3)) <> "" Then
            For x = 0 To Fg2.Rows - 1
                If Val(Fg2.TextMatrix(x, 3)) = Val(xRs("id")) Then
                    MsgBox "La cuenta seleccionada ya esta agregado", vbInformation, "Item seleccionado"
                    Exit Sub
                End If
            Next
        End If
    
        Fg2.TextMatrix(Fg2.Row, 3) = xRs("id")
        Fg2.TextMatrix(Fg2.Row, 1) = NulosC(xRs("cuenta"))
        Fg2.TextMatrix(Fg2.Row, 2) = NulosC(xRs("descripcion"))
        
        If Trim(Fg2.TextMatrix(Fg2.Row, 2)) <> "" And Trim(Fg2.TextMatrix(Fg2.Row, 3)) <> "" Then
            If Trim(Fg2.TextMatrix(Fg2.Rows - 1, 3)) <> "" Then
                Fg2.AddItem ""
                Fg2.Row = Fg2.Rows - 1: Fg2.Col = 1
            End If
        End If
    End If
    ModifiTamanioColGridCta
    Set xform = Nothing
    Set xRs = Nothing
End Sub

'--CODIGO PARA EXPORTAR A EXCEL
Private Sub FormatoExcel2()
    With Oleapp
        '--CONFIGURAR ANCHO DE LAS COLUMNAS
        .Columns("B:B").Select
        .Selection.ColumnWidth = 19.57
        .Range("A1").Select
        
        .Columns("C:C").Select
        .Selection.ColumnWidth = 11.71
        .Range("A1").Select
        
        .Columns("D:D").Select
        .Selection.ColumnWidth = 12
        .Range("A1").Select
        
        .Columns("E:E").Select
        .Selection.ColumnWidth = 11.14
        .Range("A1").Select
        
        .Columns("F:F").Select
        .Selection.ColumnWidth = 24.14
        .Range("A1").Select
        '----------------------------------
        
        '--CONFIGURAR ALTO DEL ENCABEZ 2
        .Rows("8:8").Select
        .Selection.RowHeight = 50.25
        .Range("A1").Select
        '--------------------------------
    
        .Cells(1, 2) = NomEmp
        .Cells(2, 2) = "N° R.U.C.: " & NumRUC
        .Cells(1, 7) = Date
        .Cells(4, 2) = "Libro Caja y Bancos"
        .Cells(5, 2) = "Detalle de los Movimientos del Efectivo"
        '--COMBINAR CELDAS DE LOS TITULOS
        .Range("B4:G4").Select
        With .Selection
            .HorizontalAlignment = -4108
            .VerticalAlignment = -4107
            .WrapText = False
            .Orientation = 0
            .ShrinkToFit = False
            .MergeCells = False
        End With
        .Selection.Merge
        .Selection.Font.Bold = True
        .Range("B5:G5").Select
        With .Selection
            .HorizontalAlignment = -4108
            .VerticalAlignment = -4107
            .WrapText = False
            .Orientation = 0
            .ShrinkToFit = False
            .MergeCells = False
        End With
        .Selection.Merge
        .Selection.Font.Bold = True
        .Range("A1").Select
        '----------------------------------
        
        '--SELECCIONAR TODAS LAS CELDAS DEL Y PONER TAMAÑO DE LA LETRA A 9
        .Cells.Select
        With .Selection.Font
            .Name = "Arial"
            .Size = 9
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = -4142
            .ColorIndex = -4105
        End With
        .Range("A1").Select
        '---------------------------------------------------------------------
    End With
End Sub

Private Sub FormatoExcel(pRango As String)
    With Oleapp
        If pRango <> "" Then
            '--CONFIGURAR A NEGRITA LAS CELDAS SELECCIONADAS
            .Range(pRango).Select
            .Selection.Font.Bold = True
            '----------------------------------------------
        End If
    End With
End Sub

Private Sub ExportExcel()
    Dim fs As Variant, vNumFilaFixed As Integer
    Dim i_row As Long, i_col As Long
    Dim NFILA As Long, NCOLUMN As Long, vDatTemp1 As String, vDatTemp2 As String
    Dim vNumTemp1 As Integer, vNumTemp2 As Integer, vRango1 As String
        
    BAND_INTERRUMPIR = False
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set Oleapp = CreateObject("excel.application")
    Oleapp.Visible = True
    vNumFilaFixed = Fg1.FixedRows
    With Oleapp
        .WindowState = 1
        .Workbooks.Add
        .Sheets(1).Select
        .Sheets(1).Name = "Libro1"
        
        '--PONE LOS EN ENCABEZADOS
        NFILA = 7
        NCOLUMN = 2 'COLUMNA INICIO PARA EXCEL
        For i_row = 0 To Fg1.FixedRows - 1
            NCOLUMN = 2
            vNumTemp1 = 0
            FraProgreso.Visible = True
            LblTituloProg.Caption = "Procesando exportación a excel..."
            PgBar.Max = Fg1.Rows - 1
            PgBar.Value = 0
            For i_col = Fg1.FixedCols To Fg1.Cols - 1
                If Fg1.ColWidth(i_col) > 0 Then
                    Fg1.TextMatrix(0, 1) = ""
                    Fg1.TextMatrix(0, 2) = ""
                    Fg1.TextMatrix(0, 3) = ""
                    Fg1.TextMatrix(0, 5) = ""
                    Fg1.TextMatrix(0, 7) = ""
                    Fg1.TextMatrix(0, 9) = ""
                    
                    .Cells(NFILA, NCOLUMN) = Trim(Fg1.TextMatrix(i_row, i_col))
                    vRango1 = .Cells(NFILA, NCOLUMN).Address
                    FormatoExcel vRango1
'                    If vNumTemp1 = 0 Then
'                        .Cells(NFILA, NCOLUMN) = Trim(Fg1.TextMatrix(i_row, i_col))
'                        vRango1 = .Cells(NFILA, NCOLUMN).Address
'                        FormatoExcel vRango1
'                    End If
'
'                    '--VERIFICAR SI HAY DATOS REPETIDOS EN LA FILA
'                    If Fg1.MergeRow(i_row) = True Then
'                        vDatTemp1 = Trim(Fg1.TextMatrix(i_row, i_col))
'                        If i_col + 1 < Fg1.Cols - 1 Then
'                            vDatTemp2 = Trim(Fg1.TextMatrix(i_row, i_col + 1))
'                        End If
'
'                        If vDatTemp1 = vDatTemp2 Then
'                            vNumTemp1 = 1
'                            vRango1 = .Cells(NFILA, NCOLUMN).Address
'                            FormatoExcel vRango1
'                        Else
'                            vNumTemp1 = 0
'                        End If
'                    End If
                    NCOLUMN = NCOLUMN + 1
                End If
            Next
            NFILA = NFILA + 1
        Next
        FormatoExcel ""
        
        'LLENAR LOS DATOS DEL DETALLE DE LA GRILLA
        NFILA = 9
        NCOLUMN = 2 'COLUMNA INICIO PARA EXCEL
        For i_row = vNumFilaFixed To Fg1.Rows - 1
            DoEvents
            If BAND_INTERRUMPIR = True Then
                FraProgreso.Visible = False
                Exit Sub
            End If
            NCOLUMN = 2
            vNumTemp1 = 0
            For i_col = Fg1.FixedCols To Fg1.Cols - 1
                If Fg1.ColWidth(i_col) > 0 Then
                    If vNumTemp1 = 0 Then
                        If i_col = 2 Then
                            .Cells(NFILA, NCOLUMN) = "'" & Fg1.TextMatrix(i_row, i_col)
                        Else
                            .Cells(NFILA, NCOLUMN) = Fg1.TextMatrix(i_row, i_col)
                        End If
                    End If
                    
                    '--VERIFICAR SI HAY DATOS REPETIDOS EN LA FILA
                    If Fg1.MergeRow(i_row) = True Then
                        vDatTemp1 = Trim(Fg1.TextMatrix(i_row, i_col))
                        If i_col + 1 < Fg1.Cols - 1 Then
                            vDatTemp2 = Trim(Fg1.TextMatrix(i_row, i_col + 1))
                        End If
                        
                        If vDatTemp1 = vDatTemp2 Then
                            vNumTemp1 = 1
                            vRango1 = .Cells(NFILA, NCOLUMN).Address
                            FormatoExcel vRango1
                        Else
                            vNumTemp1 = 0
                        End If
                    End If
                    
                    NCOLUMN = NCOLUMN + 1
                End If
            Next
            NFILA = NFILA + 1
            If PgBar.Value < PgBar.Max Then
                PgBar.Value = PgBar.Value + 1
            End If
        Next
        FormatoExcel2
        FraProgreso.Visible = False
        Oleapp.WindowState = 1
        .ActiveWindow.Zoom = 100
    End With
    UnirCeldasEncabezado
    Set Oleapp = Nothing   ' la aplicación; después libera la referenci
    Set fs = Nothing
    MsgBox "Los datos han sido exportados correctamente", vbInformation, "Aviso"
End Sub
'--FIN CODIGO PARA EXPORTA A EXCEL

Private Sub formatGridAlCambiarMoneda()
    If OptSol.Value = True Then
        Fg1.ColWidth(5) = 1500
        Fg1.ColWidth(6) = 1500
        Fg1.ColWidth(7) = 0
        Fg1.ColWidth(8) = 0
    Else
        Fg1.ColWidth(5) = 0
        Fg1.ColWidth(6) = 0
        Fg1.ColWidth(7) = 1500
        Fg1.ColWidth(8) = 1500
    End If
End Sub

Private Sub InsertarQuitar(pIndexBoton As Long)
    Select Case pIndexBoton
        Case 45 'INSERTAR REGI
            On Error Resume Next
            If Fg2.TextMatrix(1, 1) <> "" And Trim(Fg2.TextMatrix(Fg2.Rows - 1, 1)) <> "" Then
                Fg2.AddItem ""
                Fg2.Row = Fg2.Rows - 1: Fg2.Col = 2
            End If
        Case 46 'SUPRIMIR/DELETE
'            Fg2.TextMatrix(Fg2.Row, 1) = ""
'            Fg2.TextMatrix(Fg2.Row, 2) = ""
'            Fg2.TextMatrix(Fg2.Row, 3) = ""
            '-----
            If Fg2.Rows - 1 >= 1 Then
                Fg2.RemoveItem Fg2.Row
                Fg2.Col = 2: Fg2.Row = Fg2.Rows - 1
            Else
                LimpiarGridCta
            End If
    End Select
    ModifiTamanioColGridCta
End Sub

Private Sub LimpiarGridCta()
    Fg2.Clear
    Fg2.Rows = 1
    Fg2.FormatString = vFormatStrCuentas
    Fg2.ColWidth(3) = 0
    
    ''''''
    Fg2.ColComboList(1) = "|..."
    Fg2.Editable = flexEDKbdMouse
    Fg2.SelectionMode = flexSelectionFree
    
    Fg2.ColComboList(2) = "|..."
    Fg2.Editable = flexEDKbdMouse
    Fg2.SelectionMode = flexSelectionFree
End Sub

Private Sub fFiltro1()
    Dim i_row As Long
    vStrFiltro1 = ""
    
    For i_row = 0 To Fg2.Rows - 1
        If Trim(Fg2.TextMatrix(i_row, 3)) <> "" Then
            vStrFiltro1 = vStrFiltro1 & Fg2.TextMatrix(i_row, 3) & ", "
        End If
    Next
    If vStrFiltro1 <> "" Then
        vStrFiltro1 = Mid(vStrFiltro1, 1, Len(Trim(vStrFiltro1)) - 1)
        vStrFiltro1 = " AND con_planctas.id IN (" & vStrFiltro1 & ")"
    End If
End Sub

Sub UnirCeldas(pFila As Long, pColRang1 As Integer, pColRang2 As Integer, pCadena As String)
    With Fg1
        .MergeCells = flexMergeFree
        .Row = pFila
        
        .MergeRow(pFila) = True
'        .MergeCol(-1) = True
        .Select pFila, pColRang1, pFila, pColRang2
        '.CellAlignment = flexAlignCenterCenter
        .CellAlignment = flexAlignLeftCenter
        .Cell(flexcpText, pFila, pColRang1, pFila, pColRang2) = pCadena
    End With
End Sub
Sub UnirCeldasEncabezado()
    With Fg1
        '--UNIR LA CELDA DEL NUM CORRELATVO
        .MergeCells = flexMergeFree
        '.MergeRow(0) = True
        .MergeCol(1) = True
        .Select 0, 1, 1, 1  '
        .CellAlignment = flexAlignCenterCenter
        .Cell(flexcpText, 0, 1, 1, 1) = Fg1.TextMatrix(1, 1)
        
        '--UNIR CELDA DE FECHA
        .MergeCells = flexMergeFree
        '.MergeRow(0) = True
        .MergeCol(2) = True
        .Select 0, 2, 1, 2  '
        .CellAlignment = flexAlignCenterCenter
        .Cell(flexcpText, 0, 2, 1, 2) = Fg1.TextMatrix(1, 2)
        
        '--UNIR CELDA DESC DE LA OPERAC
        .MergeCells = flexMergeFree
        '.MergeRow(0) = True
        .MergeCol(3) = True
        .Select 0, 3, 1, 3  '
        .CellAlignment = flexAlignCenterCenter
        .Cell(flexcpText, 0, 3, 1, 3) = Fg1.TextMatrix(1, 3)
        
        '0
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
'        .MergeCol(-1) = True
        .Select 0, 4, 0, 5  '
        .CellAlignment = flexAlignCenterCenter
        .Cell(flexcpText, 0, 4, 0, 5) = "Cuenta contable asociada"  '
        '1
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .Select 0, 6, 0, 7  '
        .CellAlignment = flexAlignCenterCenter
        .Cell(flexcpText, 0, 6, 0, 7) = "Saldo y Movimiento S/."  '
        '3
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .Select 0, 8, 0, 9  '
        .CellAlignment = flexAlignCenterCenter
        .Cell(flexcpText, 0, 8, 0, 9) = "Saldo y Movimiento $"  '
    End With
End Sub
Sub LimpiarGrid()
    Fg1.Clear
    Fg1.Rows = 3
    Fg1.FormatString = vFormatString
    'If ChkSol.Value = 1 And ChkDol.Value = 0 Then
    If OptSol.Value = True Then
        Fg1.ColWidth(6) = 1500 '
        Fg1.ColWidth(7) = 1500 '
        Fg1.ColWidth(8) = 0 '
        Fg1.ColWidth(9) = 0 '
    'ElseIf ChkSol.Value = 0 And ChkDol.Value = 1 Then
    ElseIf OptDol.Value = True Then
        Fg1.ColWidth(6) = 0  '
        Fg1.ColWidth(7) = 0  '
        Fg1.ColWidth(8) = 1500  '
        Fg1.ColWidth(9) = 1500  '
    End If
    UnirCeldasEncabezado
End Sub


Private Sub CmdCerrar_Click()
    Unload Me
End Sub

Private Sub CmdConsultar_Click()
    
End Sub

Private Sub CmdImprimir_Click()
    

End Sub

Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 1 Then
        BuscarCta
    ElseIf Col = 2 Then
        BuscarCta
    End If
End Sub

Private Sub Fg2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 45  'INSERTAR REGI
            InsertarQuitar 45
        Case 46 'SUPRIMIR/DELETE
            InsertarQuitar 46
    End Select
End Sub

Private Sub Fg2_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    KeyAscii = 0
'    Select Case Col
'        Case Is <> 2
'            KeyAscii = 0
'    End Select
End Sub

Private Sub Fg2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'    If Button = 2 Then
'        PopupMenu mnumenu1
'    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        '--interrumpir
        BAND_INTERRUMPIR = True
    End If
End Sub

Private Sub Form_Load()
    vFormatString = Fg1.FormatString
    vFormatStrCuentas = Fg2.FormatString
    
    LimpiarGrid
    UnirCeldasEncabezado
    TxtFec1.Valor = "01/01/" & CStr(Year(Date)): TxtFec2.Valor = Date
    
    LimpiarGridCta
    
    FraProgreso.Left = (Me.Width - FraProgreso.Width) \ 2
    FraProgreso.Top = (Me.Height - FraProgreso.Height) \ 2
End Sub

Private Sub mnu_insertcta_Click()
    '--INSERTAR CUENTA
    InsertarQuitar 45
End Sub

Private Sub mnuquitarcta_Click()
    '--QUITAR CUENTA
    InsertarQuitar 46
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 '--BUSCAR
            BAND_INTERRUMPIR = False
            Dim vSumTotDebSol As Double, vSumTotDebDol As Double, vSumTotHabSol As Double, vSumTotHabDol As Double
            LimpiarGrid
            
            Dim RsCons As New ADODB.Recordset, RsConsPlanCta As New ADODB.Recordset
            Dim vTempRow As Long, vContador As Integer
            
            vStrCons = "SELECT id, cuenta, descripcion FROM con_planctas WHERE cuenta LIKE '10%' ORDER BY cuenta"
            RST_Busq RsConsPlanCta, vStrCons, xCon
            
            fFiltro1
            RsCons.CursorLocation = adUseClient
            If OptSol.Value = True Then
                vStrCons = "SELECT format(con_diario.idmes, '00') & '-' & con_diario.numasi as numcorr, con_cajabanco.tipope, con_cajabanco.fchope, con_cajabanco.numdoc, con_planctas.cuenta, con_planctas.descripcion, con_tc.impcom, IIf([con_cajabanco].[idmon]=2,[con_diario].[impdebdol]*[con_tc].[impcom],[con_diario].[impdebsol]) AS impdebsol, IIf([con_cajabanco].[idmon]=2,[con_diario].[imphabdol]*[con_tc].[impcom],[con_diario].[imphabsol]) AS imphabsol " _
                    & " FROM con_planctas RIGHT JOIN ((con_diario RIGHT JOIN con_cajabanco ON con_diario.idmov = con_cajabanco.id) LEFT JOIN con_tc ON con_cajabanco.fchope = con_tc.fecha) ON con_planctas.id = con_diario.idcue" _
                    & " Where con_cajabanco.tipope = 1 AND con_planctas.cuenta Like '10-[1,2]%' AND con_planctas.tipo <> 1 And con_diario.idlib = 6 AND con_diario.fchasi BETWEEN DATEVALUE('" & Trim(TxtFec1.Valor) & "') AND DATEVALUE('" & Trim(TxtFec2.Valor) & "')"
            Else
                vStrCons = "SELECT format(con_diario.idmes, '00') & '-' & con_diario.numasi as numcorr, con_cajabanco.tipope, con_cajabanco.fchope, con_cajabanco.numdoc, con_planctas.cuenta, con_planctas.descripcion, con_tc.impcom, IIf([con_cajabanco].[idmon]=1,[con_diario].[impdebsol]/[con_tc].[impcom],[con_diario].[impdebdol]) AS impdebdol, IIf([con_cajabanco].[idmon]=1,[con_diario].[imphabsol]/[con_tc].[impcom],[con_diario].[imphabdol]) AS imphabdol " _
                    & " FROM con_planctas RIGHT JOIN ((con_diario RIGHT JOIN con_cajabanco ON con_diario.idmov = con_cajabanco.id) LEFT JOIN con_tc ON con_cajabanco.fchope = con_tc.fecha) ON con_planctas.id = con_diario.idcue" _
                    & " Where con_cajabanco.tipope = 1 AND con_planctas.cuenta Like '10-[1,2]%' AND con_planctas.tipo <> 1 And con_diario.idlib = 6 AND con_diario.fchasi BETWEEN DATEVALUE('" & Trim(TxtFec1.Valor) & "') AND DATEVALUE('" & Trim(TxtFec2.Valor) & "')"
            End If
            
            vStrCons = vStrCons & vStrFiltro1
                                    
            vStrCons = vStrCons & " ORDER BY con_cajabanco.fchope"
            RST_Busq RsCons, vStrCons, xCon
            
            If RsConsPlanCta.RecordCount > 0 Then
                FraProgreso.Visible = True
                LblTituloProg.Caption = "Procesando: Detalle de los Movimientos del Efectivo"
                PgBar.Max = RsConsPlanCta.RecordCount
                PgBar.Value = 0
                RsConsPlanCta.MoveFirst
                Do While Not RsConsPlanCta.EOF
                    DoEvents
                    If BAND_INTERRUMPIR = True Then
                        FraProgreso.Visible = False
                        Exit Sub
                    End If
                    vSumTotDebSol = 0: vSumTotDebDol = 0: vSumTotHabSol = 0: vSumTotHabDol = 0
        '            If Fg1.TextMatrix(2, 1) <> "" Then Fg1.AddItem ""
                    vTempRow = Fg1.Rows - 1
                    RsCons.Filter = "cuenta = '" + Trim(RsConsPlanCta("cuenta")) + "'"
                    With RsCons
                        If .RecordCount > 0 Then
                            If vContador = 0 Then
                                Fg1.AddItem ""
                                vContador = vContador + 1
                            End If
                            .MoveFirst
                            Do While Not .EOF
                                If Fg1.TextMatrix(3, 2) <> "" Then Fg1.AddItem ""
                                Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(.Fields("numcorr"))
                                Fg1.TextMatrix(Fg1.Rows - 1, 2) = Format(Trim(.Fields("fchope")), "dd/mm/yy") '
                                Fg1.TextMatrix(Fg1.Rows - 1, 3) = Trim(NulosC(.Fields("numdoc"))) '
                                Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(.Fields("cuenta")) '
                                Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(.Fields("descripcion")) '
                                If OptSol.Value = True Then
                                    Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosN(.Fields("impdebsol")), "#,###0.00") '
                                    Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(.Fields("imphabsol")), "#,###0.00") '
                                Else
                                    Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(NulosN(.Fields("impdebdol")), "#,###0.00") '
                                    Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(NulosN(.Fields("imphabdol")), "#,###0.00") '
                                End If
                                
                                vSumTotDebSol = vSumTotDebSol + Val(Format(Fg1.TextMatrix(Fg1.Rows - 1, 6), "#####0.00")) '
                                vSumTotDebDol = vSumTotDebDol + Val(Format(Fg1.TextMatrix(Fg1.Rows - 1, 8), "#####0.00")) '
                                vSumTotHabSol = vSumTotHabSol + Val(Format(Fg1.TextMatrix(Fg1.Rows - 1, 7), "#####0.00")) '
                                vSumTotHabDol = vSumTotHabDol + Val(Format(Fg1.TextMatrix(Fg1.Rows - 1, 9), "#####0.00")) '
                                .MoveNext
                            Loop
                            Fg1.AddItem ""
                            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(vSumTotDebSol, "#,###0.00") '
                            Fg1.TextMatrix(Fg1.Rows - 1, 8) = Format(vSumTotDebDol, "#,###0.00") '
                            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(vSumTotHabSol, "#,###0.00") '
                            Fg1.TextMatrix(Fg1.Rows - 1, 9) = Format(vSumTotHabDol, "#,###0.00") '
                            Fg1.Row = Fg1.Rows - 1: Fg1.Col = 6 '
                            Fg1.CellFontBold = True
                            Fg1.Row = Fg1.Rows - 1: Fg1.Col = 7 '
                            Fg1.CellFontBold = True
                            Fg1.Row = Fg1.Rows - 1: Fg1.Col = 8 '
                            Fg1.CellFontBold = True
                            Fg1.Row = Fg1.Rows - 1: Fg1.Col = 9 '
                            Fg1.CellFontBold = True
                        End If
                    End With
                    If RsCons.RecordCount > 0 Then
                        Fg1.Col = 1: Fg1.Row = vTempRow
'                        Fg1.CellBackColor = &H8000000F
                        Fg1.Col = 2: Fg1.Row = vTempRow
'                        Fg1.CellBackColor = &H8000000F
                        UnirCeldas vTempRow, 1, 9, "Cuenta: " & NulosC(RsConsPlanCta("cuenta")) & " - " & NulosC(RsConsPlanCta("descripcion")) '
                        Fg1.CellFontBold = True
                        Fg1.AddItem ""
                    End If
                    If PgBar.Value < PgBar.Max Then
                        PgBar.Value = PgBar.Value + 1
                    End If
                    RsCons.Filter = adFilterNone
                    RsConsPlanCta.MoveNext
                Loop
                FraProgreso.Visible = False
            End If
            
            If Fg1.Rows = 3 Then
                MsgBox "No se encontraron registros...!", vbInformation, "Mensaje...!"
            End If
        Case 2 '--IMPRIMIR
            If Fg1.TextMatrix(2, 1) = "" Then
                MsgBox "No hay datos para imprimir...!", vbInformation, xTitulo
                Exit Sub
            End If
            
            On Error GoTo error
            Dim X_PRINT As New SGI2_funciones.formularios
        
            Me.MousePointer = vbHourglass
            'If MsgBox("Desea conservar el formato de la consulta", vbQuestion + vbYesNo, "Imprimir...") = vbNo Then Configurar_Grilla False
        
            X_PRINT.Imprimir_x_VSFlexGrid Fg1, "Reporte de Libro Caja y Bancos", "Detalle de los Movimientos del Efectivo", "", False, True
            Set X_PRINT = Nothing
            Me.MousePointer = vbDefault
            Exit Sub
error:
            Me.MousePointer = vbDefault
            SHOW_ERROR
        Case 3 '--EXPORTAR
            ExportExcel
        Case 4
            Unload Me
    End Select
End Sub
