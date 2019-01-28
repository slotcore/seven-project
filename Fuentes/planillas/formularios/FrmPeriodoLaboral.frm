VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmPeriodoLaboral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planillas - Periodo Laboral"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11445
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "[ Períodos Laborales]"
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
      Height          =   3810
      Left            =   45
      TabIndex        =   9
      Top             =   615
      Width           =   11280
      Begin VSFlex7Ctl.VSFlexGrid Fg2 
         Height          =   3465
         Left            =   90
         TabIndex        =   10
         Top             =   255
         Width           =   11115
         _cx             =   19606
         _cy             =   6112
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
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmPeriodoLaboral.frx":0000
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
      Height          =   525
      Index           =   0
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   10740
      Begin VB.TextBox txt 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Index           =   0
         Left            =   9510
         MaxLength       =   40
         TabIndex        =   6
         Text            =   "txt(0)"
         Top             =   60
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00800000&
         X1              =   90
         X2              =   10725
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   195
         Index           =   0
         Left            =   8925
         TabIndex        =   7
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
         TabIndex        =   3
         Top             =   105
         Width           =   990
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   3
         X1              =   -15
         X2              =   11000
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line lin 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   2
         X1              =   -30
         X2              =   11000
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   3
         X1              =   10725
         X2              =   10725
         Y1              =   -60
         Y2              =   855
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
      TabIndex        =   1
      Top             =   4500
      Width           =   10740
      Begin VB.CommandButton CmdPerLab 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   3
         Left            =   9195
         TabIndex        =   8
         Top             =   105
         Width           =   1395
      End
      Begin VB.CommandButton CmdPerLab 
         Caption         =   "&Grabar"
         Height          =   375
         Index           =   2
         Left            =   7725
         TabIndex        =   5
         Top             =   105
         Width           =   1395
      End
      Begin VB.CommandButton CmdPerLab 
         Caption         =   "&Agregar"
         Height          =   375
         Index           =   0
         Left            =   105
         TabIndex        =   0
         Top             =   105
         Width           =   1395
      End
      Begin VB.CommandButton CmdPerLab 
         Caption         =   "&Eliminar"
         Height          =   375
         Index           =   1
         Left            =   1530
         TabIndex        =   4
         Top             =   105
         Width           =   1395
      End
      Begin VB.Line lin 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   0
         X1              =   -15
         X2              =   11000
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line lin 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   -30
         X2              =   11000
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   10725
         X2              =   10725
         Y1              =   0
         Y2              =   985
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
Attribute VB_Name = "FrmPeriodoLaboral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Quehace As Integer
Dim mCorrelativo As Long
Dim mIdEmpleado As Long
Dim Agregando As Boolean
Dim SeEjecuto  As Boolean

Private Sub pConfigurarGrilla()
    With Fg2
        .Cols = 8
        .ColWidth(1) = 200
        .FixedRows = 1
        .RowHeight(0) = 500
        .ColWidth(1) = 0:
        .TextMatrix(0, 2) = "Categoría":                                    .ColWidth(2) = 2600:    .ColAlignment(2) = flexAlignLeftCenter:         .Row = 0: .Col = 2: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 3) = "Tipo Convenio" + vbCr + "(Solo Modalidad Formativa)":    .ColWidth(3) = 2700:     .ColAlignment(3) = flexAlignLeftCenter:         .Row = 0: .Col = 3: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 4) = "Fch.Inicio" + vbCr + "o Reinicio":        .ColWidth(4) = 1100:     .ColAlignment(4) = flexAlignCenterCenter:       .Row = 0: .Col = 4: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 5) = "Fch.Fin, Cese" + vbCr + "/ Suspensión":    .ColWidth(5) = 1100:    .ColAlignment(5) = flexAlignLeftCenter:         .Row = 0: .Col = 5: .CellAlignment = flexAlignCenterCenter
        .TextMatrix(0, 6) = "Tipo de Extinción del Contrato" + vbCr + "(No Considerar Modalidad Formativa)":   .ColWidth(6) = 2700:    .ColAlignment(6) = flexAlignLeftCenter:         .Row = 0: .Col = 6: .CellAlignment = flexAlignCenterCenter
        .ColEditMask(4) = "##/##/####"
        .ColEditMask(5) = "##/##/####"
        .SelectionMode = flexSelectionFree
    End With
    '*****************************************
    GRID_COMBOLIST Fg2, 7
    'Fg2.ColDataType(7) = flexDTBoolean
    '--COMBOLIST CON VSFLEXGRID
    Dim RstTmp As New ADODB.Recordset
    Dim tFormat$
    '--categoria
    RST_Busq RstTmp, "SELECT mae_categoria.id, mae_categoria.descripcion FROM mae_categoria  ;", xCon
    tFormat = Fg2.BuildComboList(RstTmp, "descripcion", "id", vbYellow)
    Fg2.ColComboList(2) = tFormat
    Set RstTmp = Nothing
    '--tipo convenio
    RST_Busq RstTmp, "SELECT mae_tipomodformativa.id, mae_tipomodformativa.descripcion FROM mae_tipomodformativa ORDER BY mae_tipomodformativa.descripcion;", xCon
    tFormat = Fg2.BuildComboList(RstTmp, "descripcion", "id", vbYellow)
    Fg2.ColComboList(3) = tFormat
    Set RstTmp = Nothing
    '--tipo extincion del contrato
    RST_Busq RstTmp, "SELECT mae_finperiodo.id, mae_finperiodo.descripcion FROM mae_finperiodo ORDER BY mae_finperiodo.descripcion;", xCon
    tFormat = Fg2.BuildComboList(RstTmp, "descripcion", "id", vbYellow)
    Fg2.ColComboList(6) = tFormat
    Set RstTmp = Nothing
    
    '*****************************************
        
    DoEvents
End Sub

Private Sub CmdPerLab_Click(Index As Integer)
    Select Case Index
        Case 0 '--AGREGAR
            pRegistroAdd
        Case 1 '--ELIMINAR
            pRegistroDel
        Case 2 '--GRABAR
            If Grabar() = False Then Exit Sub
            Unload Me
        Case 3 '--CANCELAR
            Unload Me
    End Select
End Sub




Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Quehace = 3 Then Exit Sub
    If Col <> 7 Then Exit Sub
    
    Select Case NulosN(Fg2.Cell(flexcpText, Row, 2))
        Case 1 '--trabajador
            FrmTrabajador.pRecibeLink 1
            FrmTrabajador.Show 1
        Case 2 '--pensionista
            FrmPensionista.pRecibeLink 1
            FrmPensionista.Show 1
        Case 3 '--prestador de servicio modalidad  formativa
            FrmModalidadFormativa.pRecibeLink 1
            FrmModalidadFormativa.Show 1
        Case 4 '--prestador de servicio 4ta cat
            Frm4taCategoria.Show 1
        Case 5 '--personal de tercero
            FrmPersonalTerceros.pRecibeLink 1
            FrmPersonalTerceros.Show 1
        Case Else
            
    End Select
End Sub

Private Sub fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    If Agregando = True Then Exit Sub
    If Row <= 0 Then Exit Sub
    Select Case Col
        Case 2 '--categoria
            If Fg2.Cell(flexcpText, Row, Col) = "" Then
                Fg2.TextMatrix(Row, 3) = ""
                Fg2.TextMatrix(Row, 6) = ""
                Exit Sub
            End If
            If Fg2.Cell(flexcpText, Row, Col) = 3 Then '--categoria
                Fg2.TextMatrix(Row, 6) = ""
            Else
                Fg2.TextMatrix(Row, 3) = ""
            End If
        Case 3 '--tipo convenio
            If Fg2.Cell(flexcpText, Row, 2) = "" Then
                MsgBox "Seleccione la Categoría", vbExclamation, xTitulo
                Fg2.TextMatrix(Row, 3) = ""
                Fg2.TextMatrix(Row, 6) = ""
                Fg2.Col = 2
                Fg2.SetFocus
                Exit Sub
            End If
            If Fg2.Cell(flexcpText, Row, 2) <> 3 Then '--categoria modalidad formativa
                Fg2.TextMatrix(Row, 3) = ""
            End If
            
        Case 4 '--fecha inicio
            If Fg2.TextMatrix(Row, Col) = "  /  /    " Then Exit Sub
            If IsDate(Fg2.TextMatrix(Row, Col)) = False Then
                MsgBox "La Fecha de Inicio o Reinicio es incorrecta", vbExclamation, xTitulo
                Fg2.TextMatrix(Row, Col) = ""
                Exit Sub
            End If
            Fg2.TextMatrix(Row, Col) = Format(Fg2.TextMatrix(Row, Col), "dd/mm/yyyy")
        Case 5 '--fecha fin
            If Fg2.TextMatrix(Row, Col) = "  /  /    " Then Exit Sub
            If IsDate(Fg2.TextMatrix(Row, Col)) = False Then
                MsgBox "La Fecha de Fin, Cese / Suspención  es incorrecta", vbExclamation, xTitulo
                Fg2.TextMatrix(Row, Col) = ""
                Exit Sub
            End If
            
            If IsDate(Fg2.TextMatrix(Row, 4)) = True Then
                If CDate(Fg2.TextMatrix(Row, 4)) > CDate(Fg2.TextMatrix(Row, 5)) Then
                    MsgBox "La Fecha de Fin es inferior a la Fecha de Inicio", vbExclamation, xTitulo
                    Fg2.TextMatrix(Row, Col) = ""
                    Exit Sub
                End If
            End If
            
            Fg2.TextMatrix(Row, Col) = Format(Fg2.TextMatrix(Row, Col), "dd/mm/yyyy")
        Case 6 '--tipo de extincion del contrato
            If Fg2.Cell(flexcpText, Row, 2) = "" Then
                MsgBox "Seleccione la Categoría", vbExclamation, xTitulo
                Fg2.TextMatrix(Row, 3) = ""
                Fg2.TextMatrix(Row, 6) = ""
                Fg2.Col = 2
                Fg2.SetFocus
                Exit Sub
            End If
            If Fg2.Cell(flexcpText, Row, 2) = 3 Then '--categoria modalidad formativa
                Fg2.TextMatrix(Row, 6) = ""
            End If
    
    End Select
    Exit Sub
error:
    SHOW_ERROR Me.Name, "fg_CellChanged (" + CStr(Index) + ")"
End Sub

Private Sub fg2_EnterCell()
    If Fg2.Row < 1 Then Exit Sub
        
    If Fg2.Col = 3 Then
        If Fg2.Cell(flexcpText, Fg2.Row, 2) = 3 Then  '--categoria modalidad formativa
            Fg2.Editable = flexEDKbdMouse
        Else
            Fg2.Editable = flexEDNone
        End If
        
    ElseIf Fg2.Col = 6 Then
        If Fg2.Cell(flexcpText, Fg2.Row, 2) = 3 Or Trim(Fg2.Cell(flexcpText, Fg2.Row, 2)) = "" Then '--categoria modalidad formativa
            Fg2.Editable = flexEDNone
        Else
            Fg2.Editable = flexEDKbdMouse
        End If
    Else
        Fg2.Editable = flexEDKbdMouse
    End If
    
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
    If Fg2.Rows = 1 Then
        Quehace = 1
    Else
        Quehace = 2
        Fg2.Row = 1
        Fg2.Col = 2
        Fg2.SetFocus
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

    Agregando = True

    '************************************************************
       
    nSQL = "SELECT pla_periodolaboral.* From pla_periodolaboral Where (((pla_periodolaboral.idemp) = " & mIdEmpleado & ")) ORDER BY pla_periodolaboral.fchini ASC ; "

    RST_Busq RstTmp, nSQL, xCon
    
    Fg2.Rows = 1
    
    If RstTmp.RecordCount <> 0 Then RstTmp.MoveFirst
    Do While Not RstTmp.EOF
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(Fg2.Rows - 1, 1) = NulosN(RstTmp("corr"))
        Fg2.TextMatrix(Fg2.Rows - 1, 2) = NulosC(RstTmp(("idcat")))
        If NulosC(RstTmp("idmodfor")) <> 0 Then
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = NulosC(RstTmp("idmodfor"))
        End If
        Fg2.TextMatrix(Fg2.Rows - 1, 4) = NulosC(RstTmp("fchini"))
        Fg2.TextMatrix(Fg2.Rows - 1, 5) = NulosC(RstTmp("fchfin"))
        If NulosC(RstTmp("idfinper")) <> 0 Then
            Fg2.TextMatrix(Fg2.Rows - 1, 6) = NulosC(RstTmp("idfinper"))
        End If
        RstTmp.MoveNext
    Loop
    
    Set RstTmp = Nothing
    Exit Sub
error:
    Agregando = False
    Set RstTmp = Nothing
    habilitar CmdPerLab, False
    CmdPerLab(3).Enabled = True
    SHOW_ERROR Me.Name, "pPonerDatos"
End Sub


Function Grabar() As Boolean

    If fValidarDatos() = False Then Exit Function
    
    If MsgBox("Seguro desea " + IIf(Quehace = 1, "Grabar", "Modificar") + " los datos del Periodo Laboral", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Function

    Dim RstDet As New ADODB.Recordset
    Dim xId As Integer
    Dim A&
'    On Error GoTo LaCague

    xCon.BeginTrans

    '*****************************************************
    If Quehace = 1 Then
    
    Else
        xCon.Execute "Delete from pla_periodolaboral where idemp =  " & mIdEmpleado & " ;"
    End If
    RST_Busq RstDet, "SELECT TOP 1 * FROM pla_periodolaboral ; ", xCon
    '************************************************************TAB 0
   
    For A = 1 To Fg2.Rows - 1
        RstDet.AddNew
        RstDet("idemp") = mIdEmpleado
        RstDet("corr") = A
        RstDet("idcat") = Fg2.Cell(flexcpText, A, 2)
        If NulosN(Fg2.Cell(flexcpText, A, 2)) = 3 Then '--categoria -modalidad formativa
            RstDet("idmodfor") = Fg2.TextMatrix(A, 3)   '--tipo convenio - modalidad formatva
        Else
            RstDet("idfinper") = NulosN(Fg2.Cell(flexcpText, A, 6)) '--tipo de extincion del contrato
        End If
        If IsDate(Fg2.TextMatrix(A, 4)) = True Then
            RstDet("fchini") = CDate(Fg2.TextMatrix(A, 4))
        End If
        If IsDate(Fg2.Cell(flexcpText, A, 5)) = True Then
            RstDet("fchfin") = CDate(Fg2.Cell(flexcpText, A, 5))
        End If
        RstDet.Update
    Next A
        
    MsgBox "Los datos del Periodo Laboral se " + IIf(Quehace = 1, "grabaron", "modificaron") + " con éxito", vbInformation, xTitulo

    xCon.CommitTrans
    Set RstDet = Nothing
    Grabar = True
    Exit Function
LaCague:
    Set RstDet = Nothing
    xCon.RollbackTrans
    MsgBox "No se pudo guardar los datos del Periodo Laboral por el siguiente motivo: " + vbCr + Trim(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Function

Private Function fValidarDatos() As Boolean
    
    Dim band As Integer
    
    If Fg2.Rows = 1 Then
        MsgBox "No hay registros", vbExclamation, xTitulo
        CmdPerLab(0).SetFocus
        Exit Function
    End If
    
    
    '--VALIDAR EL INGRESO DE LOS DATOS
    Dim mRow  As Long
    Dim mCol As Long '--COLUMNA A POSICIONAR SI FALTAN DATOS
    mCol = -1
    For mRow = 1 To Fg2.Rows - 1
        If NulosN(Fg2.Cell(flexcpText, mRow, 2)) = 0 Then '--categoria
            MsgBox "Falta especificar la Categoría", vbExclamation, xTitulo
            mCol = 2:          Exit For
        End If
        If NulosN(Fg2.Cell(flexcpText, mRow, 2)) = 3 And NulosN(Fg2.Cell(flexcpText, mRow, 3)) = 0 Then '--categoria - modalidad formativa
            MsgBox "Falta especificar el tipo de convenio de Modalidad Formativa", vbExclamation, xTitulo
            mCol = 3:          Exit For
        End If
        If IsDate(Fg2.TextMatrix(mRow, 4)) = False Then
            MsgBox "Falta especificar la Fecha de Inicio o Reinicio", vbExclamation, xTitulo
            mCol = 4:          Exit For
        End If
        If mRow < Fg2.Rows - 1 Then '--si la fila actual es inferior al total de filas => obligar que se ingrese los datos
            If IsDate(Fg2.TextMatrix(mRow, 5)) = False Then
                MsgBox "Falta especificar la Fecha de Fin, Cese / Suspensión", vbExclamation, xTitulo
                mCol = 5:          Exit For
            End If
            If NulosN(Fg2.Cell(flexcpText, mRow, 2)) <> 3 And NulosN(Fg2.Cell(flexcpText, mRow, 6)) = 0 Then '--categoria - modalidad formativa
                MsgBox "Falta especificar el tipo de Extinción del Contrato", vbExclamation, xTitulo
                mCol = 6:          Exit For
            End If
        End If
    Next mRow
    If mCol <> -1 Then
        Agregando = True:  Fg2.Row = mRow: Fg2.Col = mCol: Agregando = False
        Fg2.SetFocus
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
    If Quehace = 3 Then Exit Sub
    Agregando = True
    If Fg2.Rows > 1 Then
        If NulosN(Fg2.Cell(flexcpText, Fg2.Rows - 1, 2)) = 0 Then  '--categoria
            MsgBox "Falta ingresar la Categoría", vbExclamation, xTitulo
            mCol = 2
        Else
            If NulosN(Fg2.Cell(flexcpText, Fg2.Rows - 1, 2)) = 3 And NulosN(Fg2.Cell(flexcpText, Fg2.Rows - 1, 3)) = 0 Then    '--categoria modalidad formativa
                    MsgBox "Falta ingresar el Tipo de Convenio Modalidad Formativa", vbExclamation, xTitulo
                    mCol = 3
            End If
        End If
        If mCol = 0 Then
            If IsDate(Fg2.TextMatrix(Fg2.Rows - 1, 4)) = False Then
                MsgBox "Falta ingresar la Fecha de Inicio", vbExclamation, xTitulo
                mCol = 4
            ElseIf IsDate(Fg2.TextMatrix(Fg2.Rows - 1, 5)) = False Then
                MsgBox "Falta ingresar la Fecha de Cese", vbExclamation, xTitulo
                mCol = 5
            ElseIf NulosN(Fg2.Cell(flexcpText, Fg2.Rows - 1, 2)) <> 3 And NulosN(Fg2.Cell(flexcpText, Fg2.Rows - 1, 6)) = 0 Then
                    MsgBox "Falta ingresar el Tipo de Extinción del Contrato", vbExclamation, xTitulo
                    mCol = 6
            Else

                Fg2.AddItem ""
                mCol = 2
            End If
        End If
    Else
        Fg2.AddItem ""
        mCol = 2
    End If
    Fg2.Row = Fg2.Rows - 1
    Fg2.Col = mCol
    Fg2.SetFocus
    Agregando = False
End Sub

Private Sub pRegistroDel()
    If Fg2.Rows = 1 Then Exit Sub
    If Fg2.Row < 1 Then Exit Sub
    If MsgBox("Seguro desea eliminar el Registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    Fg2.RemoveItem Fg2.Row
    
End Sub

Private Sub fg2_KeyUp(KeyCode As Integer, Shift As Integer)
    If Quehace = 3 Then Exit Sub
    If KeyCode = 45 Then
        pRegistroAdd
    End If
    If KeyCode = 46 Then
        pRegistroDel
    End If
End Sub


