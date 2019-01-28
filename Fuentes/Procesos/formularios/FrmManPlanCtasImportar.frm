VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmManPlanCtasImportar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Importar Plan de Cuentas"
   ClientHeight    =   7200
   ClientLeft      =   105
   ClientTop       =   600
   ClientWidth     =   11805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   11805
   Begin MSComDlg.CommonDialog cmm 
      Left            =   2565
      Top             =   990
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   6840
      Left            =   60
      TabIndex        =   0
      Top             =   375
      Width           =   11760
      _cx             =   20743
      _cy             =   12065
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
      FormatString    =   $"FrmManPlanCtasImportar.frx":0000
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
            Object.ToolTipText     =   "MSExcel"
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Crear Formato a Importar"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Importar"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cerrar"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6585
         Top             =   15
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483637
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManPlanCtasImportar.frx":003C
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManPlanCtasImportar.frx":0490
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManPlanCtasImportar.frx":05FC
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManPlanCtasImportar.frx":0B44
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManPlanCtasImportar.frx":0EDC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManPlanCtasImportar.frx":0FEE
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmManPlanCtasImportar.frx":1100
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "Importar"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "Grabar"
      End
      Begin VB.Menu Menu1_4 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu Menu1_5 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_6 
         Caption         =   "Formato para Importar"
      End
   End
End
Attribute VB_Name = "FrmManPlanCtasImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--
'--DE LA IMPRESION
Dim T_RPT_PERIODO As String '--PERIODO DEL REPORTE
Dim T_RPT_TITULO As String  '--TITULO DE REPORTE

Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim Agregando As Boolean

Dim RstCta As New ADODB.Recordset

'------------

Private Sub pImprimir()
    On Error GoTo error
    Dim oPrint As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    oPrint.Imprimir_x_VSFlexGrid Fg1, "Plan de Cuentas", "", " ", False, True
    Set oPrint = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pImprimir"
End Sub

Private Sub Fg1_EnterCell()
    If Fg1.Col = Fg1.Cols - 1 Then
        Fg1.Editable = flexEDNone
    Else
        If Fg1.Col < 2 Or Fg1.Row < 1 Then
            Fg1.Editable = flexEDNone
            Exit Sub
        End If
        Fg1.Editable = flexEDKbdMouse
    End If
End Sub


Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If QueHace = 3 Then Exit Sub
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim nTitulo As String

    Select Case Col
        Case 4, 5 '--CUENTA DESTINO 3=DEBE, 4=HABER
            ReDim xCampos(2, 4) As String
            xCampos(0, 0) = "N° Cuenta":    xCampos(0, 1) = "nombre":        xCampos(0, 2) = "1500":    xCampos(0, 3) = "C"
            xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "5000":    xCampos(1, 3) = "C"

'            nSQL = "SELECT con_planctas.cuenta , con_planctas.descripcion , con_planctas.id " _
'                + vbCr + " FROM con_planctas ORDER BY con_planctas.cuenta ASC "
                
            If Col = 4 Then
                nTitulo = "Buscando Cta Destino [Debe]"
            Else
                nTitulo = "Buscando Cta Destino [Haber]"
            End If
            
            nSQL = ""
            RstCta.Filter = ""
            CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio, , RstCta
            
        Case 6, 7, 8 '--
            ReDim xCampos(1, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
                
            nSQL = "SELECT con_planctasdes.id, con_planctasdes.descripcion , con_planctasdes.descripcion AS nombre " _
                + vbCr + " FROM con_planctasdes " _
                + vbCr + " ORDER BY con_planctasdes.descripcion;"
        
        If Col = 7 Then
            nTitulo = "Buscando 1ra. Distribución"
        ElseIf Col = 8 Then
            nTitulo = "Buscando 2da. Distribución"
        Else
            nTitulo = "Buscando 3ra. Distribución"
        End If
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "descripcion", "descripcion", Principio, ""
            
    End Select


    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    Agregando = True
    
    Fg1.TextMatrix(Row, Col) = xRs.Fields("nombre") & ""
    Fg1.TextMatrix(Row, Col + 9) = xRs.Fields("id") & ""
    
    Agregando = False
    Set xRs = Nothing
    Exit Sub
SALIR:
    Set xRs = Nothing
    Agregando = False
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    If Agregando = True Then Exit Sub
    If Row < 2 Then Exit Sub
    Select Case Col
        Case 2, 3
            RstCta.Filter = "id=" & NulosN(Fg1.TextMatrix(Row, 1))
            If RstCta.RecordCount <> 0 Then
                RstCta.Fields("nombre") = NulosC(Fg1.TextMatrix(Row, 2))
                RstCta.Fields("descripcion") = NulosC(Fg1.TextMatrix(Row, 3))
                RstCta.Update
            End If
        Case 9, 11
            If Trim(Fg1.TextMatrix(Row, Col + 1)) = "-1" Then
                Fg1.TextMatrix(Row, Col + 1) = ""
            End If
        Case 10, 12
            If Trim(Fg1.TextMatrix(Row, Col - 1)) = "-1" Then
                Fg1.TextMatrix(Row, Col - 1) = ""
            End If
        Case 4 To 8
            If Trim(Fg1.TextMatrix(Row, Col)) = "" Then
                Fg1.TextMatrix(Row, Col + 9) = ""
            End If
    End Select
    Exit Sub
error:
    SHOW_ERROR Me.Name, "Fg1_CellChanged"
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Row = 0 Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub



Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo error
    If KeyCode = 114 Or KeyCode = vbKeyInsert Then 'F3 = Agregar Item
        pRegistroAdd
    End If
    If KeyCode = 115 Or KeyCode = vbKeyDelete Then
        pRegistroDel 0   'Eliminar Item
    End If
    Exit Sub
error:
    SHOW_ERROR Me.Name, "Fg1_KeyUp"
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace = 3 Then
'            PopupMenu Menu4
        Else
            PopupMenu Menu1
        End If
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
    
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo error
    SeEjecuto = False
    Agregando = False
    QueHace = 1
    
    CentrarFrm Me
    Fg1.Rows = 2
    pConfigurarGrilla
    
    
    Exit Sub
error:
    SHOW_ERROR
End Sub


Private Sub pConfigurarGrilla(Optional F_CONSERVAR_FORMATO As Boolean = False)
    '--PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA
    
    If F_CONSERVAR_FORMATO = True Then Fg1.Clear
    
    Fg1.FrozenCols = 0
    Agregando = True
    With Fg1
        .Cols = 18
        .FixedRows = 2
        .ColWidth(0) = 200
        .FrozenCols = 3
        .RowHeight(0) = 500
        .RowHeight(1) = 250
        UNIR_CELDAS Fg1, 0, 1, 0, 3, " ", flexAlignCenterCenter
        UNIR_CELDAS Fg1, 0, 4, 0, 5, "Transferencia Automática", flexAlignCenterCenter
        UNIR_CELDAS Fg1, 0, 6, 0, 8, "Estado de Resultados", flexAlignCenterCenter
        UNIR_CELDAS Fg1, 0, 9, 0, 10, "Destino del Saldo", flexAlignCenterCenter
        UNIR_CELDAS Fg1, 0, 11, 0, 12, "Distribuir la Cuenta en Balance" + vbCr + "General en función al saldo", flexAlignCenterCenter
        '--DE LOS ID'S
        UNIR_CELDAS Fg1, 0, 13, 0, 17, "DE LOS ID'S", flexAlignCenterCenter
        '--DATOS DE FILA
        .TextMatrix(1, 1) = "Id":                   .ColWidth(1) = 500:     .ColAlignment(1) = flexAlignRightCenter:     .Row = 0: .Col = 1: .CellAlignment = flexAlignRightCenter
        .TextMatrix(1, 2) = "Nº Cuenta":            .ColWidth(2) = 1100:    .ColAlignment(2) = flexAlignLeftBottom:     .Row = 0: .Col = 2: .CellAlignment = flexAlignLeftCenter
        .TextMatrix(1, 3) = "Nombre de la Cuenta":  .ColWidth(3) = 3500:    .ColAlignment(3) = flexAlignLeftBottom:     .Row = 0: .Col = 3: .CellAlignment = flexAlignLeftCenter
        
        .TextMatrix(1, 4) = "Cta Destino Debe":     .ColWidth(4) = 1500:    .ColAlignment(4) = flexAlignLeftBottom:     .Row = 0: .Col = 4: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 5) = "Cta Destino Haber":    .ColWidth(5) = 1500:    .ColAlignment(5) = flexAlignLeftBottom:     .Row = 0: .Col = 5: .CellAlignment = flexAlignCenterCenter
        '----------------------
        .TextMatrix(1, 6) = "1ra Distribución":     .ColWidth(6) = 1300:    .ColAlignment(6) = flexAlignLeftBottom:     .Row = 0: .Col = 6: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 7) = "2da Distribución":     .ColWidth(7) = 1300:    .ColAlignment(7) = flexAlignLeftBottom:     .Row = 0: .Col = 7: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 8) = "3ra Distribución":     .ColWidth(8) = 0:       .ColAlignment(8) = flexAlignLeftBottom:     .Row = 0: .Col = 8: .CellAlignment = flexAlignCenterCenter
        
        .TextMatrix(1, 9) = "Debe(D)":              .ColWidth(9) = 900:     .ColAlignment(9) = flexAlignCenterCenter:     .Row = 0: .Col = 9: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 10) = "Haber(H)":            .ColWidth(10) = 900:   .ColAlignment(10) = flexAlignCenterCenter:     .Row = 0: .Col = 10: .CellAlignment = flexAlignCenterCenter
        
        .ColDataType(9) = flexDTBoolean:            .ColDataType(10) = flexDTBoolean
        
        .TextMatrix(1, 11) = "Si":                  .ColWidth(11) = 1150:   .ColAlignment(11) = flexAlignCenterCenter:     .Row = 0: .Col = 11: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(1, 12) = "No":                  .ColWidth(12) = 1150:   .ColAlignment(12) = flexAlignCenterCenter:     .Row = 0: .Col = 12: .CellAlignment = flexAlignCenterCenter
        
        .ColDataType(11) = flexDTBoolean:           .ColDataType(12) = flexDTBoolean
        
        .TextMatrix(1, 13) = "Id Cta Destino Debe":   .ColWidth(13) = 500:    .ColAlignment(13) = flexAlignRightBottom
        .TextMatrix(1, 14) = "Id Cta Destino Haber":  .ColWidth(14) = 500:    .ColAlignment(14) = flexAlignRightBottom
        .TextMatrix(1, 15) = "Id 1ra Distribución":   .ColWidth(15) = 500:    .ColAlignment(15) = flexAlignRightBottom
        .TextMatrix(1, 16) = "Id 2da Distribución":   .ColWidth(16) = 500:    .ColAlignment(16) = flexAlignRightBottom
        .TextMatrix(1, 17) = "Id 3ra Distribución":   .ColWidth(17) = 500:    .ColAlignment(17) = flexAlignRightBottom
        
        GRID_COMBOLIST Fg1, 4
        GRID_COMBOLIST Fg1, 5
        GRID_COMBOLIST Fg1, 6
        GRID_COMBOLIST Fg1, 7
        GRID_COMBOLIST Fg1, 8

        OCULTAR_COL Fg1, 13, 17
        
    End With
    DoEvents
    Agregando = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RstCta = Nothing
End Sub

Private Sub Menu1_1_Click()
    pImportar
End Sub

Private Sub Menu1_4_Click()
    pImprimir
End Sub

Private Sub Menu1_6_Click()
    pCrearFormato
End Sub

'----

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 2 Then Grabar
    If Button.Index = 3 Then pImprimir
    If Button.Index = 5 Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then pCrearFormato
    If ButtonMenu.Index = 3 Then pImportar
End Sub


Private Sub pExportar()
    On Error GoTo error
    Dim oExport As New SGI2_funciones.formularios

    oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg1, "Plan de Cuentas", "", , "Plan de Cuentas"
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportar"
End Sub

Private Sub pImportar()
    Dim vPath As String
    Dim mRow&, mCol&
    cmm.FileName = ""
    cmm.CancelError = False
    'cmm.Filter = "Archivos xls (*.csv)|*.csv"
    cmm.Filter = "Archivos xls (*.xls)|*.xls"
    cmm.ShowOpen
    vPath = cmm.FileName
    If vPath = "" Then Exit Sub
    If ArchivoExiste(vPath) = False Then
        MsgBox "El archivo no Existe", vbExclamation, xTitulo
        Exit Sub
    End If
    Set RstCta = Nothing
    '--definir el recordset temporal
    RstCta.Fields.Append "id", adBigInt, -1
    RstCta.Fields.Append "nombre", adVarChar, 50
    RstCta.Fields.Append "descripcion", adVarChar, 250
    RstCta.Open
    Fg1.Rows = 2
    
    '**********************************************************
    On Error GoTo error
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    
    Me.MousePointer = vbHourglass
    
    objExcel.Visible = False
    objExcel.SheetsInNewWorkbook = 1
    'Crea el Libro
    objExcel.Workbooks.Open vPath

    Dim xFila&
    xFila = 4
    Agregando = True
    With objExcel.ActiveSheet
        Do While NulosN(.Cells(xFila, 1)) <> 0
            DoEvents
            Fg1.Rows = Fg1.Rows + 1
            
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = xFila - 3
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(Replace(.Cells(xFila, 2), "'", ""))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(.Cells(xFila, 3))
            Select Case Mid(Fg1.TextMatrix(Fg1.Rows - 1, 2), 1, 1)
                Case 1, 2, 3, 7 '--Debe
                    Fg1.TextMatrix(Fg1.Rows - 1, 9) = -1
                Case 4, 5, 6, 8, 9 '--haber
                    Fg1.TextMatrix(Fg1.Rows - 1, 10) = -1
            End Select
            
            '--agregando al recordset
            RstCta.AddNew
            RstCta.Fields("id") = xFila - 3
            RstCta.Fields("nombre") = NulosC(Fg1.TextMatrix(Fg1.Rows - 1, 2))
            RstCta.Fields("descripcion") = NulosC(Fg1.TextMatrix(Fg1.Rows - 1, 3))
            RstCta.Update
                        
            xFila = xFila + 1
        Loop
    End With
    
    Set objExcel = Nothing
    Me.MousePointer = vbDefault
    
    MsgBox "El Archivo se Importó Correctamente", vbInformation, xTitulo
    '**********************************************************
    '--colocando grupos
    GRID_AGRUPAR Fg1, 1
    Agregando = False
    Exit Sub
error:
    Me.MousePointer = vbDefault
    Set objExcel = Nothing
    Agregando = False
    SHOW_ERROR Me.Name, "pImportar"
End Sub

Private Sub pCrearFormato()

    On Error GoTo error
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    objExcel.SheetsInNewWorkbook = 1
    objExcel.WindowState = 1
    objExcel.Workbooks.Add
    
    With objExcel.ActiveSheet
        .Cells(1, 1) = "Plan de Cuentas"
        .Cells(3, 1) = "Id"
        .Cells(3, 2) = "Nº. Cuenta"
        .Cells(3, 3) = "Descripción"
        '---------
        .Columns(1).ColumnWidth = 5
        .Columns(2).ColumnWidth = 11
        .Columns(3).ColumnWidth = 60
        '---
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 15
        .Cells(3, 1).Font.Bold = True
        .Cells(3, 2).Font.Bold = True
        .Cells(3, 3).Font.Bold = True
    End With
    MsgBox "Proceda a ingresar la información según los Parámetros Solicitados" + vbCr + "Luego proceda a Importar...", vbInformation, xTitulo
    objExcel.Visible = True
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
error:
    Set objExcel = Nothing
End Sub


Function Grabar() As Boolean
    '--OBS SE HAN MODIFICADO LAS RELACIONES DE LAS SIG. TABLAS CON con_planctas
    '--desactivo la opcion EXIGIR INTEGRIDAD REFERENCIAL
    '--mae_documentocta
    '--alm_inventario
    '--con_bancocuenta
    '--con_origen
    '--con_destino
    '--con_diario
    '--con_balancedet
    '--con_estadosdet
    '--MODIFICADO AL 02-01-08

    If fValidarDatos() = False Then Exit Function
    If MsgBox("Seguro desea grabar el Listado de Cuentas" + vbCr + "Al grabar el listado, eliminará las cuentas registradas" + vbCr + "Desea continuar?", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Function

    Dim RstCab As New ADODB.Recordset
    Dim xId&
    Dim mRow&
    On Error GoTo LaCague
    
    xCon.BeginTrans
    Me.MousePointer = vbHourglass
    RST_Busq RstCab, "SELECT TOP 1 * FROM con_planctas", xCon
    xCon.Execute "delete from con_planctas"
    For mRow = Fg1.FixedRows To Fg1.Rows - 1
        DoEvents
'        xId = HallaCodigoTabla("con_planctas", xCon, "id")
        
        xId = NulosN(Fg1.TextMatrix(mRow, 1))
        
        RstCab.AddNew
        RstCab("id") = xId
        
        RstCab("cuenta") = NulosC(Fg1.TextMatrix(mRow, 2))
        RstCab("descripcion") = NulosC(Fg1.TextMatrix(mRow, 3))
        
        RstCab("ctadesdeb") = NulosN(Fg1.TextMatrix(mRow, 13))
        RstCab("ctadeshab") = NulosN(Fg1.TextMatrix(mRow, 14))
        RstCab("iddes") = NulosN(Fg1.TextMatrix(mRow, 15))
        RstCab("iddes2") = NulosN(Fg1.TextMatrix(mRow, 16))
        RstCab("iddes3") = NulosN(Fg1.TextMatrix(mRow, 17))
       
        RstCab("tipsal") = IIf(Trim(Fg1.TextMatrix(mRow, 9)) = "-1", "D", IIf(Trim(Fg1.TextMatrix(mRow, 10)) = "-1", "H", ""))
        RstCab("dissegsal") = IIf(Trim(Fg1.TextMatrix(mRow, 11)) = "-1", "-1", "0")
        
        RstCab.Update
    Next mRow
    '---ver si dependen de otras cuentas
    
    '-----DEL TIPO   1 = cuentas; 0 = registro
    '--SI DEPENDE DE OTRA CUENTA
    Dim xRs As New ADODB.Recordset
    Dim NumCta As String
    Dim mPos&
    Dim mTipo&
    For mRow = Fg1.FixedRows To Fg1.Rows - 1
        xId = NulosN(Fg1.TextMatrix(mRow, 1))
        NumCta = StrReverse(NulosC(Fg1.TextMatrix(mRow, 1)))
        mPos = InStr(NumCta, "-")
        If mPos <> 0 Then
            NumCta = StrReverse(Mid(NumCta, mPos + 1))
            RST_Busq xRs, "SELECT id FROM con_planctas WHERE (((cuenta)= '" & NumCta & "'));", xCon
            If xRs.EOF = False Or xRs.BOF = False Or xRs.RecordCount <> 0 Then
                xCon.Execute "UPDATE con_planctas SET tipo = 1 WHERE id = " & CStr(xRs.Fields("id")) '--CUENTA
            End If
        End If
        '--SI TIENEN CUENTAS QUE DEPENDEN DE ESTE
        Set xRs = Nothing
        NumCta = NulosC(Fg1.TextMatrix(mRow, 2))
        RST_Busq xRs, "SELECT con_planctas.id, con_planctas.cuenta FROM con_planctas WHERE (((con_planctas.id)<>" & CStr(xId) & ") AND ((con_planctas.cuenta) Like '" & NumCta & "%'));", xCon
        If xRs.EOF = False Or xRs.BOF = False Or xRs.RecordCount <> 0 Then
            mTipo = 1 '--ES CUENTA
        Else
            mTipo = 0 '--ES REGISTRO
        End If
        '-----
        xCon.Execute "update con_planctas set tipo=" & mTipo & " where id = " & xId
        '-----
    Next mRow
    
    MsgBox "La Lista de Cuentas se grabó con éxito", vbInformation, xTitulo
    xCon.CommitTrans
    Grabar = True
    Me.MousePointer = vbDefault
    Set RstCab = Nothing:
    Exit Function
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing:       Set xRs = Nothing
    Me.MousePointer = vbDefault
    MsgBox "No se pudo guardar la lista de Cuentas Cntables por el siguiente motivo: " + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
End Function

Private Function fValidarDatos() As Boolean
    '--VALIDAR QUE EL REGISTRO NO ESTE REGISTRADO
    '-----
    Dim mRow&
    
    With Fg1
        For mRow = .FixedRows + 1 To .Rows - 1
            If NulosC(.TextMatrix(mRow, 2)) = "" Or NulosC(.TextMatrix(mRow, 3)) = "" Then
                MsgBox "Falta Completar los Datos de N°.Cta o Descripción", vbExclamation, xTitulo
                Agregando = True:  .Row = mRow: .Col = IIf(Trim(.TextMatrix(mRow, 2)) = "", 1, 2): Agregando = False
                
                Exit Function

            End If
        Next mRow
    End With
    
    If Fg1.FixedRows = Fg1.Rows Then
        MsgBox "No hay información para grabar", vbExclamation, xTitulo
        Exit Function
    End If

    fValidarDatos = True
End Function
 

Private Sub pRegistroDel(Index As Integer)
    Dim mRow&, mRowDel&
    Dim mIdCta&
    If QueHace = 3 Then Exit Sub
    If Fg1.Row <= 1 Then
        MsgBox "La fila no se puede eliminar" + vbCr + "Seleccione una fila correcta", vbExclamation, xTitulo
        Exit Sub
    End If
    '--eliminar registro
    
    mRowDel = Fg1.Row
    mIdCta = Fg1.TextMatrix(mRowDel, 1)
    '--buscar si existe relacion alguna de las demas cuentas con la cuenta a eliminar
    Dim fEncuentra As Boolean
    Me.MousePointer = vbHourglass
    For mRow = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(mRow, 13)) = mIdCta Then
            fEncuentra = True
        ElseIf NulosN(Fg1.TextMatrix(mRow, 14)) = mIdCta Then
            fEncuentra = True
        End If
        If fEncuentra = True Then
            Exit For
        End If
    Next mRow
    If fEncuentra = True Then
        If MsgBox("La cuenta que Desea eliminar se encuentra Relacionado a" + vbCr + "N°.Cta: " & Fg1.TextMatrix(mRow, 2) + vbCr + "Descripción: " + Fg1.TextMatrix(mRow, 3) + vbCr + "Si desea continuar se anularán aquellos registros que tengan alguna relación con la cuenta a eliminar", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Me.MousePointer = vbDefault
            Exit Sub
        End If
    Else
        If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo, xTitulo) = vbNo Then
            Me.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    '--eliminar de la grilla
    Fg1.RemoveItem mRowDel
    '--eliminar cuenta del recodset
    RstCta.Filter = "id=" & mIdCta
    If RstCta.RecordCount <> 0 Then
        RstCta.Delete
        RstCta.Update
    End If

    '--limpiar los registros que contengan a la cuenta seleccionada
    Agregando = True
    If fEncuentra = True Then
        For mRow = 1 To Fg1.Rows - 1
            If NulosN(Fg1.TextMatrix(mRow, 13)) = mIdCta Then
                Fg1.TextMatrix(mRow, 4) = ""
                Fg1.TextMatrix(mRow, 13) = ""
            ElseIf NulosN(Fg1.TextMatrix(mRow, 14)) = mIdCta Then
                Fg1.TextMatrix(mRow, 5) = ""
                Fg1.TextMatrix(mRow, 15) = ""
            End If
        Next mRow
    End If
    
    '--colocando grupos
    Agregando = True
    GRID_AGRUPAR Fg1, 1
    Agregando = False
    '---------
    If Fg1.FixedRows = Fg1.Rows - 1 Then
        Fg1.Row = 2
    Else
        Fg1.Row = mRowDel - 1
    End If
    Fg1.Col = 3

    Me.MousePointer = vbDefault
End Sub

Private Sub pRegistroAdd()
    If Fg1.TextMatrix(Fg1.Rows - 1, 2) = "" And Fg1.TextMatrix(Fg1.Rows - 1, 3) = "" Then
        Exit Sub
    End If
    Fg1.AddItem ""
    Fg1.Row = Fg1.Rows - 1
    Fg1.Col = 1
    Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosN(Fg1.TextMatrix(Fg1.Rows - 2, 1)) + 1
    '--add cuenta
    RstCta.AddNew
    RstCta.Fields("id") = Fg1.TextMatrix(Fg1.Rows - 1, 1)
    RstCta.Update
    
    '--colocando grupos
    Agregando = True
    GRID_AGRUPAR Fg1, 1
    Agregando = False
    '---------
    
    Fg1.SetFocus
            
End Sub

