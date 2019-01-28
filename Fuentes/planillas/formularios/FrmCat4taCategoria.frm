VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmCat4taCategoria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prestador de Servicio -    4ta Categoría"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "[ Solicitudes de Suspención ]"
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
      Height          =   2145
      Left            =   45
      TabIndex        =   11
      Top             =   975
      Width           =   7650
      Begin VSFlex7Ctl.VSFlexGrid Fg1 
         Height          =   1815
         Left            =   150
         TabIndex        =   12
         Top             =   270
         Width           =   7410
         _cx             =   13070
         _cy             =   3201
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmCat4taCategoria.frx":0000
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
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      Height          =   900
      Index           =   0
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   7650
      Begin VB.TextBox txt 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Index           =   0
         Left            =   6225
         MaxLength       =   40
         TabIndex        =   8
         Text            =   "txt(0)"
         Top             =   60
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   870
         MaxLength       =   11
         TabIndex        =   0
         Text            =   "txt(1)"
         Top             =   510
         Width           =   1440
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00800000&
         X1              =   90
         X2              =   7485
         Y1              =   435
         Y2              =   435
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   195
         Index           =   0
         Left            =   5640
         TabIndex        =   9
         Top             =   150
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbl_persona 
         AutoSize        =   -1  'True
         Caption         =   "lbl_persona"
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
         Height          =   195
         Left            =   105
         TabIndex        =   5
         Top             =   105
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RUC"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   4
         Top             =   600
         Width           =   345
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   3
         X1              =   -15
         X2              =   8500
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line lin 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   2
         X1              =   -30
         X2              =   8500
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   3
         X1              =   7635
         X2              =   7635
         Y1              =   0
         Y2              =   915
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   15
         X2              =   30
         Y1              =   0
         Y2              =   990
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      Height          =   585
      Index           =   12
      Left            =   45
      TabIndex        =   2
      Top             =   3135
      Width           =   7650
      Begin VB.CommandButton Cmd 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   5895
         TabIndex        =   10
         Top             =   105
         Width           =   1395
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   2
         Left            =   4380
         TabIndex        =   7
         Top             =   105
         Width           =   1395
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "&Agregar"
         Height          =   375
         Index           =   0
         Left            =   105
         TabIndex        =   1
         Top             =   105
         Width           =   1395
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "&Eliminar"
         Height          =   375
         Index           =   1
         Left            =   1530
         TabIndex        =   6
         Top             =   105
         Width           =   1395
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   0
         X1              =   -15
         X2              =   8500
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line lin 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   -30
         X2              =   8500
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   7635
         X2              =   7635
         Y1              =   15
         Y2              =   1000
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   1000
      End
   End
End
Attribute VB_Name = "FrmCat4taCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim QueHace As Integer
Dim mCorrelativo As Long
Dim mIdEmpleado As Long
Dim Agregando As Boolean
Dim SeEjecuto  As Boolean

Private Sub pConfigurarGrilla()

    With Fg1
        .Cols = 6
        .ColWidth(1) = 200
        .FixedRows = 1
        .RowHeight(0) = 500
        .ColWidth(1) = 0:
        .TextMatrix(0, 2) = "Fecha de " + vbCr + "Presentación": .ColWidth(2) = 1000:  .ColAlignment(2) = flexAlignLeftCenter:         .Row = 0: .Col = 2: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 3) = "Número de " + vbCr + "Operación":   .ColWidth(3) = 2000:  .ColAlignment(3) = flexAlignLeftCenter:         .Row = 0: .Col = 3: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 4) = "Ejercicio":                         .ColWidth(4) = 800:   .ColAlignment(4) = flexAlignCenterCenter:        .Row = 0: .Col = 4: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 5) = "Medio de" + vbCr + "Presentación":  .ColWidth(5) = 3000:  .ColAlignment(5) = flexAlignLeftCenter:         .Row = 0: .Col = 5: .CellAlignment = flexAlignCenterCenter
        .ColEditMask(2) = "##/##/####"
        .ColEditMask(4) = "####"
        .SelectionMode = flexSelectionFree
        GRID_COMBOLIST Fg1, 2
    End With
    '*****************************************
    '--COMBOLIST CON VSFLEXGRID
    Dim RstTmp As New ADODB.Recordset
    RstTmp.Fields.Append "Nombre", adVarChar, -1
    RstTmp.Fields.Append "id", adInteger, -1
    RstTmp.Open
    RstTmp.AddNew
    RstTmp.Fields("nombre") = "Internet"
    RstTmp.Fields("id") = 1
    RstTmp.AddNew
    RstTmp.Fields("nombre") = "Dependencia SUNAT"
    RstTmp.Fields("id") = 2

    Dim tFormat$
    tFormat = Fg1.BuildComboList(RstTmp, "nombre", "id", vbYellow)
    Fg1.ColComboList(5) = tFormat
    Set RstTmp = Nothing
    '*****************************************
        
    DoEvents
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0 '--AGREGAR
            pRegistroAdd
        Case 1 '--ELIMINAR
            pRegistroDel
        Case 2 '--GRABAR
            If Grabar() = False Then Exit Sub
            FrmNomina.pCargarDatosPeriodoLaboral
            Unload Me
        Case 3 '--CANCELAR
            Unload Me
    End Select
End Sub



Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col <> 2 Then Exit Sub
    If Row < 1 Then Exit Sub
    '--invocar al formulario de fecha
    Dim obj As New SGI2_funciones.formularios
    obj.FechaSeleccionar Fg1, Row, Col, Fg1.TextMatrix(Row, Col)
    Set obj = Nothing
End Sub

Private Sub fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    If Agregando = True Then Exit Sub
    If Row <= 0 Then Exit Sub
    Select Case Col
        Case 2 '--FCH PRESENTACION
            If Fg1.TextMatrix(Row, Col) = "  /  /    " Then Exit Sub
            If IsDate(Fg1.TextMatrix(Row, Col)) = False Then
                MsgBox "La Fecha de Presentación es incorrecta", vbExclamation, xTitulo
                Fg1.TextMatrix(Fg1.Rows - 1, 2) = ""
                Exit Sub
            End If
            Fg1.TextMatrix(Row, Col) = Format(Fg1.TextMatrix(Row, Col), "dd/mm/yyyy")
        Case 4 '--EJERCICIO
            If Trim(Fg1.TextMatrix(Row, Col)) = "" Then Exit Sub
            If IsNumeric(Fg1.TextMatrix(Row, Col)) = False Then
                MsgBox "El Ejercicio es incorrecto", vbExclamation, xTitulo
                Fg1.TextMatrix(Fg1.Rows - 1, 4) = ""
            End If
    End Select
    Exit Sub
error:
    SHOW_ERROR Me.Name, "fg_CellChanged (" + CStr(Index) + ")"
End Sub



Private Sub Fg1_EnterCell()
    Fg1.Editable = flexEDKbdMouse
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
   
    SeEjecuto = False
    
    pConfigurarGrilla
    mCorrelativo = mCorr
    With FrmNomina
        mIdEmpleado = .txt(0).Text
        lbl_persona.Caption = .lbl_persona(0).Caption
    End With
    '------
    LimpiaText txt
    pPonerDatos
    If Trim(txt(1).Text) = "" Then
        QueHace = 1
    Else
        QueHace = 3
    End If
    If Fg1.Rows = 1 Then
        txt(1).SetFocus
    Else
        Fg1.Row = 1
        Fg1.Col = 2
        Fg1.SetFocus
    End If
    SeEjecuto = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    centrarfrm Me
    SeEjecuto = False
End Sub

'****************************************************************************************

Private Sub pPonerDatos()
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo error

    
    nSQL = "SELECT pla_categoria3.* From pla_categoria3 WHERE (((pla_categoria3.idemp)=" & mIdEmpleado & "));"

    RST_Busq RstTmp, nSQL, xCon
    
    If RstTmp.RecordCount = 0 Then
        
        Exit Sub
    End If
    
    Agregando = True

    txt(1).Text = NulosC(RstTmp.Fields("numruc"))

    '************************************************************
       
    nSQL = "SELECT pla_categoria3susp.* From pla_categoria3susp Where (((pla_categoria3susp.idemp) = " & mIdEmpleado & ")) ORDER BY pla_categoria3susp.fchpre; "

    RST_Busq RstTmp, nSQL, xCon
    
    Fg1.Rows = 1
    
    If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    Do While Not RstTmp.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosN(RstTmp("corr"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstTmp(("fchpre")))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(RstTmp("numope"))
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = NulosC(RstTmp("ejercicio"))
        Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosC(RstTmp("medio"))
        RstTmp.MoveNext
    Loop
    
    Set RstTmp = Nothing
    Exit Sub
error:
    Agregando = False
    Set RstTmp = Nothing
    habilitar Cmd, False
    Cmd(3).Enabled = True
    SHOW_ERROR Me.Name, "pPonerDatos"
End Sub


Function Grabar() As Boolean

    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(QueHace = 1, "Grabar", "Modificar") + " los datos del [ Prestador de Servicio - 4ta Categoría ]", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    Dim RstCab As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim xId As Integer

'    On Error GoTo LaCague

    xCon.BeginTrans

    '*****************************************************
    If QueHace = 1 Then
        RST_Busq RstCab, "SELECT TOP 1 * FROM pla_categoria3 ; ", xCon
        RstCab.AddNew
    Else
        RST_Busq RstCab, "SELECT * FROM pla_categoria3 WHERE idemp =  " & mIdEmpleado & " ;", xCon
        xCon.Execute "delete from pla_categoria3susp where idemp =  " & mIdEmpleado & " ;"
        
    End If
    RST_Busq RstDet, "SELECT TOP 1 * FROM pla_categoria3susp ; ", xCon
    RstCab("idemp") = mIdEmpleado
    RstCab("numruc") = Trim(txt(1).Text)
    '************************************************************TAB 0
    RstCab.Update
    
    For A = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("idemp") = mIdEmpleado
        RstDet("corr") = A
        RstDet("fchpre") = CDate(Fg1.TextMatrix(A, 2))
        RstDet("numope") = Mid(Fg1.TextMatrix(A, 3), 1, RstDet("numope").DefinedSize)
        RstDet("ejercicio") = Fg1.TextMatrix(A, 4)
        RstDet("medio") = Fg1.Cell(flexcpText, A, 5)
        RstDet.Update
    Next A
        
    MsgBox "Los datos del [ Prestador de Servicio - 4ta Categoría ] se " + IIf(QueHace = 1, "grabaron", "modificaron") + " con éxito", vbInformation, xTitulo

    xCon.CommitTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    Grabar = True
    Exit Function

LaCague:
    Set RstCab = Nothing
    Set RstDet = Nothing
    xCon.RollbackTrans
    MsgBox "No se pudo guardar los datos del [ Prestador de Servicio - 4ta Categoría ] por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Function

Private Function fValidarDatos() As Boolean
    
    Dim band As Integer
    
    If Trim(txt(1).Text) = "" Then
        MsgBox "Falta especificar el N°. RUC", vbExclamation, xTitulo
        Exit Function
    End If
    If Len(txt(1).Text) < 11 Then
        MsgBox "El RUC es Incorrecto", vbExclamation, xTitulo
        Exit Function
    End If
    '--VALIDAR EL INGRESO DE LOS DATOS
    Dim mRow  As Long
    Dim mCol As Long '--COLUMNA A POSICIONAR SI FALTAN DATOS
    mCol = -1
    For mRow = 1 To Fg1.Rows - 1
        If IsDate(Fg1.TextMatrix(mRow, 2)) = False Then
            MsgBox "Falta especificar la Fecha de Presentación", vbExclamation, xTitulo
            mCol = 2:          Exit For
        ElseIf IsNumeric(Fg1.TextMatrix(mRow, 4)) = False Then
            MsgBox "Falta especificar el Ejercicio: ", vbExclamation, xTitulo
            mCol = 4:          Exit For
        End If
    Next mRow
    If mCol <> -1 Then
        Agregando = True:  Fg1.Row = mRow: Fg1.Col = mCol: Agregando = False
        Fg1.SetFocus
        Exit Function
    End If
    
    fValidarDatos = True
    
End Function

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Select Case Index
        Case 1 '--
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End Select
End Sub

Private Sub pCargarGrid()
    On Error GoTo error

    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pCargarGrid"
End Sub

Private Sub pRegistroAdd()
    Dim mCol%
    If QueHace = 3 Then Exit Sub
    Agregando = True
    If Fg1.Rows > 1 Then
        If IsDate(Fg1.TextMatrix(Fg1.Rows - 1, 2)) = False Then
            MsgBox "Falta ingresar la Fecha de Presentación", vbExclamation, xTitulo
            mCol = 2
        ElseIf Trim(Fg1.TextMatrix(Fg1.Rows - 1, 3)) = "" Then
            MsgBox "Falta ingresar el N° de Operación", vbExclamation, xTitulo
            mCol = 3
        ElseIf IsNumeric(Fg1.TextMatrix(Fg1.Rows - 1, 4)) = False Then
            MsgBox "Falta ingresar el Ejercicio", vbExclamation, xTitulo
            mCol = 4
        ElseIf Trim(Fg1.TextMatrix(Fg1.Rows - 1, 5)) = "" Then
            MsgBox "Falta ingresar el Medio de Presentación", vbExclamation, xTitulo
            mCol = 5
        Else
            Fg1.AddItem ""
            mCol = 2
        End If
    Else
        Fg1.AddItem ""
        mCol = 2
    End If
    Fg1.Row = Fg1.Rows - 1
    Fg1.Col = mCol
    Fg1.SetFocus
    Agregando = False
End Sub

Private Sub pRegistroDel()
    If Fg1.Rows = 1 Then Exit Sub
    If Fg1.Row <= 1 Then Exit Sub
    If MsgBox("Seguro desea eliminar el Registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    Fg1.RemoveItem Fg1.Row
    
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 45 Then
        pRegistroAdd
    End If
    If KeyCode = 46 Then
        pRegistroDel
    End If
End Sub
