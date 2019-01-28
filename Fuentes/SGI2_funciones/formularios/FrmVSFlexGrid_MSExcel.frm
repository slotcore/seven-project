VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVSFlexGrid_MSExcel 
   BackColor       =   &H00976600&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   Icon            =   "FrmVSFlexGrid_MSExcel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   1905
      TabIndex        =   1
      Top             =   1680
      Width           =   1725
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Ejecutar"
      Height          =   375
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   1680
      Width           =   1725
   End
   Begin VB.Frame FraProgreso 
      BackColor       =   &H00976600&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   30
      TabIndex        =   14
      Top             =   315
      Width           =   5955
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   375
         Left            =   90
         TabIndex        =   15
         Top             =   105
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Index           =   1
         X1              =   5940
         X2              =   5940
         Y1              =   15
         Y2              =   675
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         Index           =   0
         X1              =   0
         X2              =   0
         Y1              =   -105
         Y2              =   555
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Index           =   1
         X1              =   30
         X2              =   5940
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   0
         X1              =   15
         X2              =   5895
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VB.Frame fraPropiedades 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   780
      TabIndex        =   2
      Top             =   2355
      Visible         =   0   'False
      Width           =   4920
      Begin VB.Frame Frame1 
         Caption         =   "Mostrar"
         Height          =   510
         Left            =   3195
         TabIndex        =   12
         Top             =   150
         Width           =   1635
         Begin VB.CheckBox ChkLeyenda 
            Caption         =   "Leyenda"
            Height          =   195
            Left            =   195
            TabIndex        =   13
            Top             =   225
            Width           =   1005
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Con Datos"
         Height          =   1215
         Left            =   105
         TabIndex        =   9
         Top             =   135
         Width           =   1515
         Begin VB.OptionButton OptconDatosDetalle1 
            Caption         =   "Detallado"
            Height          =   210
            Left            =   165
            TabIndex        =   11
            Top             =   645
            Width           =   1185
         End
         Begin VB.OptionButton OptConDatoResum1 
            Caption         =   "Resumido"
            Height          =   195
            Left            =   165
            TabIndex        =   10
            Top             =   315
            Value           =   -1  'True
            Width           =   1185
         End
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "&Cancelar"
         Height          =   345
         Index           =   1
         Left            =   3225
         TabIndex        =   8
         Top             =   1065
         Width           =   1560
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "&Aceptar"
         Height          =   345
         Index           =   0
         Left            =   3225
         TabIndex        =   7
         Top             =   690
         Width           =   1560
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Gráfico"
         Height          =   1230
         Left            =   1680
         TabIndex        =   3
         Top             =   135
         Width           =   1410
         Begin VB.OptionButton OptTipGrafCircular 
            Caption         =   "Circular"
            Height          =   195
            Left            =   165
            TabIndex        =   6
            Top             =   915
            Width           =   1020
         End
         Begin VB.OptionButton OptTipGrafLinea 
            Caption         =   "Lineas"
            Height          =   195
            Left            =   165
            TabIndex        =   5
            Top             =   615
            Width           =   1020
         End
         Begin VB.OptionButton OptTipGrafBarra1 
            Caption         =   "Barras"
            Height          =   195
            Left            =   165
            TabIndex        =   4
            Top             =   300
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.Shape shap 
         Height          =   1470
         Index           =   1
         Left            =   0
         Top             =   45
         Width           =   4860
      End
   End
   Begin VB.Label lbl 
      BackColor       =   &H00B97C00&
      Caption         =   "lbl(1)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1125
      TabIndex        =   17
      Top             =   930
      Width           =   4860
   End
   Begin VB.Label lbl 
      BackColor       =   &H00B97C00&
      Caption         =   "Exportar Datos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   60
      TabIndex        =   20
      Top             =   45
      Width           =   1530
   End
   Begin VB.Label lbl 
      BackColor       =   &H00B97C00&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   4440
      TabIndex        =   19
      Top             =   30
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label lbl 
      BackColor       =   &H00B97C00&
      Caption         =   "Exportando:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   30
      TabIndex        =   18
      Top             =   930
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackColor       =   &H00B97C00&
      Height          =   255
      Left            =   -180
      TabIndex        =   16
      Top             =   15
      Width           =   6180
   End
End
Attribute VB_Name = "FrmVSFlexGrid_MSExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'--PARA EXPORTAR A EXCEL
'''-----------
'----EXPORTAR DATOS POR VSFlexGrid
'----POR: JOHAN CASTRO
'----18/12/07
'----DESCRIPCION DE PARAMETROS A USAR EN FUNCION
'RECIBE_LINK_EXPORT(GRID1 ,T_TITULO_2 ,T_PERIODO_2 ,T_TITULO_1_2 As String ,T_NOMBRE_A_EXPORTAR_2 )
    'T_TITULO_2:: TITULO
    'T_PERIODO_2::PERIODO (OPCIONAL)
    'T_TITULO_1_2::SEGUNDO TITULO (OPCIONAL)
    'T_NOMBRE_A_EXPORTAR_2::INDICA EL NOMBRE QUE APARECERA EN LA VENTANA DE EXPORTAR EXCEL (OPCIONAL)

'----01/08/08 Johan Castro
'----agregar rutina para exportar mediante un rst temporal
'----el recordset puede tener varias columnas, se considera del array
'----los valores del array se muestra de la sig. manera
    '0::Nombre a Mostrar;
    '1::nombre de Campo del Rst;
    '2::alineacion(0::derecha, 1::centro, 2::izquierda);
    '3::ancho de columna, se tomara como si fuera el ancho de un vsflexgrid


Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE

                   
Dim M_POS_INCIAL As Long            '--INDICA LA POSICION EN EL REPORTE PARA IMPRIMIR LA INFORMACION

Dim Q_COLS&, Q_ROW&, Q_COL&, Q_COL_1&

Dim Q_COL_INICIAL As Integer        '-- INDICA LA POSICION INICIAL DE GRILLA A IMPRIMIR



Const xls_ROW_INICIO As Integer = 10
Const xls_COL_INICIO As Integer = 2

Dim SGI_JC As New SGI2_funciones.JC_Varios

'---------------------------------
Dim GRID As Object
Dim T_TITULO As String
Dim T_TITULO_1 As String
Dim T_PERIODO As String
'-----------

Dim SeEjecuto As Boolean

'---
Dim RstTmp As ADODB.Recordset
Dim ArrCampos

Public Sub RECIBE_LINK_EXPORT(GRID1 As Object, _
                            T_TITULO_2 As String, _
                            Optional T_PERIODO_2 As String = "", _
                            Optional T_TITULO_1_2 As String = "", _
                            Optional T_NOMBRE_A_EXPORTAR_2 As String = "", _
                            Optional RstTmp_2 As ADODB.Recordset, _
                            Optional xcampos_2)
                
    T_TITULO = T_TITULO_2
    T_PERIODO = T_PERIODO_2
    T_TITULO_1 = T_TITULO_1_2
    
    Me.MousePointer = vbDefault
        '---------
    If IsArray(xcampos_2) = False Then
        '--con flex grid
        Set GRID = GRID1
    Else
        '--con recordset
        Set RstTmp = New ADODB.Recordset
        Set RstTmp = RstTmp_2
        ArrCampos = xcampos_2
    End If
    
    If T_NOMBRE_A_EXPORTAR_2 = "" Then
        LBL(0).Tag = "": LBL(1).Tag = ""
    Else
        LBL(0).Tag = LBL(0).Caption: LBL(1).Tag = UCase(T_NOMBRE_A_EXPORTAR_2)
    End If
    Me.Caption = "Exportar Datos:"
    
    LBL(0).Caption = "": LBL(1).Caption = ""
    '---------
        
    DoEvents
End Sub


Private Sub cmd_Click(index As Integer)
    Select Case index
        Case 0
            PONER_DATOS
        Case 1
            Set GRID = Nothing
            Unload Me
    End Select
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    SeEjecuto = True
    If cmd(0).Enabled = True Then
        If IsArray(ArrCampos) = False Then
            PONER_DATOS
        Else
            PONER_DATOS_RST
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        '--interrumpir
        BAND_INTERRUMPIR = True
    ElseIf KeyCode = vbKeyF9 And Shift = 0 Then
        BAND_INTERRUMPIR = True
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    SGI_JC.CentrarFrm Me
    SGI_JC.FrmOcultarBoton Me.hwnd, 2

End Sub


Private Sub PONER_DATOS()
                
    On Error GoTo ERROR
    Dim GRUPO_TEXTO As String
    Dim GRUPO_CUENTA As Integer
    
    Dim Q_COL_ANTERIOR As Integer '--ES LA POSICION ANTERIOR A LA COLUMNA ACTUAL
    
    Dim xls_ROW&, xls_COL&, XCOL_ALINEACION&
    Dim N_RANGO1, N_RANGO2, N_RANGO_UNIR As String
    Dim xls_TAMANO_GRUPO As Integer
                            
                            
    
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")

    'determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1

    'abre el Libro
    objExcel.Workbooks.Add
    '------------------------------------
    Me.MousePointer = vbHourglass
    LBL(2).Visible = True
    SGI_JC.habilitar cmd, False
    BAND_INTERRUMPIR = False
    LBL(0).Caption = LBL(0).Tag:    LBL(1).Caption = LBL(1).Tag
    Me.Caption = LBL(0).Tag
    '------------------------------------
    With objExcel.ActiveSheet
        '-----------------------------------------------------------------
        PONER_EMPRESA objExcel
        '-----------------------------------------------------------------
        xls_TAMANO_GRUPO = -1
        
        xls_ROW = xls_ROW_INICIO
        xls_COL = xls_COL_INICIO
        
        '------
        PONER_ENCABEZADO objExcel, GRID

        PgBar.Value = 0
        PgBar.Max = GRID.Rows - GRID.FixedRows
        For Q_ROW = GRID.FixedRows To GRID.Rows - 1
            DoEvents
'            PgBar.Refresh
            If BAND_INTERRUMPIR = True Then GoTo salir
            xls_COL = xls_COL_INICIO
            
            GRUPO_CUENTA = 0
'''''''            '--MOVERSE A TRAVES DEL GRID
'''''''            GRID.Row = Q_ROW
            
            PgBar.Value = PgBar.Value + 1
            
            For Q_COL = 1 To GRID.Cols - 1
''                DoEvents
                If BAND_INTERRUMPIR = True Then GoTo salir
                If GRID.ColWidth(Q_COL) <> 0 And Not GRID.ColHidden(Q_COL) Then
                    '--ALINEACION
'''''                    '--MOVERSE A TRAVES DEL GRID
'''''                    GRID.Col = Q_COL
                    '--
                    If Q_COL_INICIAL = Q_COL Then
                        GRUPO_TEXTO = CStr(GRID.TextMatrix(Q_ROW, Q_COL))
                        Q_COL_ANTERIOR = Q_COL
                    Else
                        Q_COL_ANTERIOR = OBTENER_COL_ANTERIOR(CInt(Q_COL))
                    End If
                    '--
                    '--VER GRUPOS
                    If GRID.MergeCells = flexMergeFree And GRID.MergeRow(Q_ROW) = True And CStr(GRID.TextMatrix(Q_ROW, Q_COL_ANTERIOR)) = CStr(GRID.TextMatrix(Q_ROW, Q_COL)) Then
                        If GRUPO_CUENTA = 0 Then
                            GRUPO_TEXTO = CStr(GRID.TextMatrix(Q_ROW, Q_COL))
                            xls_TAMANO_GRUPO = FIN_COL_GRUPO(Q_ROW, Q_COL)
                        End If
                    ElseIf GRID.MergeCells = flexMergeFree And GRID.MergeRow(Q_ROW) = True And CStr(GRID.TextMatrix(Q_ROW, Q_COL_ANTERIOR)) <> CStr(GRID.TextMatrix(Q_ROW, Q_COL)) Then
                        If GRUPO_CUENTA = 1 Then
                            GRUPO_TEXTO = CStr(GRID.TextMatrix(Q_ROW, Q_COL))
                            xls_TAMANO_GRUPO = FIN_COL_GRUPO(Q_ROW, Q_COL)
                            GRUPO_CUENTA = 0
                        End If
                    Else
                        GRUPO_TEXTO = "xxxxxxxxx"
                        GRUPO_CUENTA = 0
                        xls_TAMANO_GRUPO = -1
                    End If
                    '--------
                    If (GRUPO_TEXTO = CStr(GRID.TextMatrix(Q_ROW, Q_COL)) Or GRUPO_TEXTO = "xxxxxxxxx") And GRUPO_CUENTA = 0 Then
                    
                                               
                        'fg1.CellFontBold = True
                        .Cells(xls_ROW, xls_COL).Font.Bold = GRID.CellFontBold
                        '--COLOR AL TEXTO

                        If xls_TAMANO_GRUPO <> -1 Then
                            '--UNIR CELDAS
                            N_RANGO1 = objExcel.Cells(xls_ROW, xls_COL).Address
                            N_RANGO2 = objExcel.Cells(xls_ROW, xls_COL + (xls_TAMANO_GRUPO - Q_COL)).Address
                            N_RANGO_UNIR = N_RANGO1 & ":" & N_RANGO2
                            UNIR_CELDA objExcel, N_RANGO_UNIR
                        End If
                        XCOL_ALINEACION = COL_ALINEACION(CInt(Q_COL))
                        .Cells(xls_ROW, xls_COL).HorizontalAlignment = XCOL_ALINEACION
                        If XCOL_ALINEACION = -4152 Then '--DERECHO
                            If (IsNumeric(GRID.TextMatrix(Q_ROW, Q_COL)) = True) Then
                                .Cells(xls_ROW, xls_COL) = NulosN(GRID.TextMatrix(Q_ROW, Q_COL))
                            Else
                                .Cells(xls_ROW, xls_COL) = "'" + GRID.TextMatrix(Q_ROW, Q_COL)
                            End If
                        ElseIf XCOL_ALINEACION = -4131 Then '--IZQUIERDO
                            If (IsNumeric(GRID.TextMatrix(Q_ROW, Q_COL)) = True) Then
                                .Cells(xls_ROW, xls_COL) = "'" + GRID.TextMatrix(Q_ROW, Q_COL)
                            Else
                                .Cells(xls_ROW, xls_COL) = GRID.TextMatrix(Q_ROW, Q_COL)
                            End If
                        Else
                            .Cells(xls_ROW, xls_COL) = "'" + GRID.TextMatrix(Q_ROW, Q_COL)
                        End If
                       
                    End If
                    
                    If GRID.MergeCells = flexMergeFree Then GRUPO_CUENTA = 1
                    
                    xls_COL = xls_COL + 1
                End If
                
            Next Q_COL
                xls_ROW = xls_ROW + 1

                        
        Next Q_ROW
        '------
    End With
salir:
    
    If BAND_INTERRUMPIR = True Then
        MsgBox "El proceso de exportación se interrumpió", vbInformation, xTitulo
    Else
''''        MsgBox "El proceso de exportación terminó con exito", vbInformation, xTitulo
    End If
    objExcel.Visible = True
    objExcel.WindowState = 1
    Set objExcel = Nothing

    Me.MousePointer = vbDefault
    Unload Me
    Exit Sub
    
ERROR:
    
    BAND_INTERRUMPIR = False
    Me.MousePointer = vbDefault
    Set objExcel = Nothing
    SGI_JC.SHOW_ERROR "", "", True, IIf(Err.Number <> 50290, "", "No manipule el archivo hasta que termine de exportar!!!!")
    LBL(0).Caption = "Presione F9 para Salir..."
    Unload Me
End Sub

Private Function COL_ALINEACION(Col As Integer, Optional ENCABEZADO As Boolean = False) As Long
    '--ESTA FUNCION DEVOLVERA LA CONSTANTE DE ALINEACION QUE SOPORTA EL EXCEL EN FUNCION A LA ALINEACION DEL GRID
    '--------------------------------------------
    '--------------------------------------------
    'xlCenter= -4108
    'xlLeft= -4131
    'xlRight = -4152

    Dim Alineacion As Integer
    Dim XVALOR As Variant
    If ENCABEZADO = True Then
        XVALOR = GRID.CellAlignment
    Else
        XVALOR = GRID.ColAlignment(Col)
    End If
    Select Case XVALOR
        Case 3, 4, 5:   Alineacion = -4108 '--CENTRADO
        Case 2, 0, 1:   Alineacion = -4131 '--IZQUIERDO
        Case 6, 7, 8:  Alineacion = -4152 '--DERECHO
        Case Else
            Alineacion = -4131
    End Select
    COL_ALINEACION = Alineacion
End Function


Private Function OBTENER_COL_ANTERIOR(Q_COL_INI As Integer) As Integer
    Dim X_POS  As Integer
    Dim N_VALOR As String
    Dim M_ZISE_GRUPO As Integer
    If X_POS = 1 Then
        OBTENER_COL_ANTERIOR = 1
        Exit Function
    End If
    For X_POS = Q_COL_INI - 1 To 1 Step -1
        If GRID.ColWidth(X_POS) <> 0 Then
            OBTENER_COL_ANTERIOR = X_POS
            Exit Function
        End If
    Next
    
End Function


Private Function FIN_COL_GRUPO(K_ROW As Long, Q_COL_INI As Long) As Integer
    '--ESTA FUNCION CALCULARA EL TAMAÑO HORIZONTAL DEL GRUPO
    Dim X_POS  As Integer
    Dim N_VALOR As String
    Dim M_COL_GRUPO_FIN As Integer
    Dim M_GRUPO As Integer
    M_COL_GRUPO_FIN = 0
    N_VALOR = CStr(GRID.TextMatrix(K_ROW, Q_COL_INI))
    For X_POS = Q_COL_INI + 1 To GRID.Cols - 1
        If GRID.ColWidth(X_POS) <> 0 Then
            If GRID.MergeCells = flexMergeFree And GRID.MergeRow(K_ROW) = True And N_VALOR = CStr(GRID.TextMatrix(K_ROW, X_POS)) Then
                M_COL_GRUPO_FIN = X_POS 'M_COL_GRUPO_FIN + 1
            Else
                Exit For
            End If
        End If
    Next
    
    If M_COL_GRUPO_FIN = 0 Then
        FIN_COL_GRUPO = -1
    Else
        FIN_COL_GRUPO = M_COL_GRUPO_FIN
    End If
    
End Function

Private Sub PONER_ENCABEZADO(objExcel As Object, GRID As Object)
  
    Dim GRUPO_TEXTO As String
    Dim GRUPO_CUENTA As Integer
    
    Dim Q_COL_ANTERIOR As Integer '--ES LA POSICION ANTERIOR A LA COLUMNA ACTUAL
    
    Dim Q_ROW1&
    
    Dim xls_ROW&, xls_COL&
    Dim N_RANGO1, N_RANGO2, N_RANGO_UNIR As String
    Dim xls_TAMANO_GRUPO As Integer

    Q_COL_INICIAL = 1
    For Q_COL = 1 To GRID.Cols - 1
        If GRID.ColWidth(Q_COL) <> 0 And Not GRID.ColHidden(Q_COL) Then
            Q_COL_INICIAL = Q_COL
            Exit For
        End If
    Next Q_COL
    
    xls_ROW = xls_ROW_INICIO - GRID.FixedRows
    If GRID.FixedRows > 0 Then
        For Q_ROW1 = 0 To GRID.FixedRows - 1
            GRUPO_CUENTA = 0
            'objExcel.Columns(xls_ROW).RowHeight = GRID.RowHeight(Q_ROW1) / 20
            'Rows("9:9").RowHeight = 18
            objExcel.Rows(xls_ROW).RowHeight = GRID.RowHeight(Q_ROW1) / 20
            xls_COL = xls_COL_INICIO
            
            '----
            For Q_COL = 1 To GRID.Cols - 1
                If GRID.ColWidth(Q_COL) <> 0 And Not GRID.ColHidden(Q_COL) Then
                    '--MOVERSE A TRAVES DEL GRID
                    GRID.Row = Q_ROW1
                    GRID.Col = Q_COL
                    '-------------------
                    If Q_COL_INICIAL = Q_COL Then
                        GRUPO_TEXTO = CStr(GRID.TextMatrix(Q_ROW1, Q_COL))
                        Q_COL_ANTERIOR = Q_COL
                    Else
                        Q_COL_ANTERIOR = OBTENER_COL_ANTERIOR(CInt(Q_COL))
                    End If
                    '--VER GRUPOS
                    If GRID.MergeCells = flexMergeFree And GRID.MergeRow(Q_ROW1) = True And CStr(GRID.TextMatrix(Q_ROW1, Q_COL_ANTERIOR)) = CStr(GRID.TextMatrix(Q_ROW1, Q_COL)) Then
                        If GRUPO_CUENTA = 0 Then
                            GRUPO_TEXTO = CStr(GRID.TextMatrix(Q_ROW1, Q_COL))
                            xls_TAMANO_GRUPO = FIN_COL_GRUPO(Q_ROW1, Q_COL)
                        End If
                    ElseIf GRID.MergeCells = flexMergeFree And GRID.MergeRow(Q_ROW1) = True And CStr(GRID.TextMatrix(Q_ROW1, Q_COL_ANTERIOR)) <> CStr(GRID.TextMatrix(Q_ROW1, Q_COL)) Then
                        If GRUPO_CUENTA = 1 Then
                            GRUPO_TEXTO = CStr(GRID.TextMatrix(Q_ROW1, Q_COL))
                            xls_TAMANO_GRUPO = FIN_COL_GRUPO(Q_ROW1, Q_COL)
                            GRUPO_CUENTA = 0
                        End If
                    Else
                        
                        GRUPO_TEXTO = "xxxxxxxxx"
                        GRUPO_CUENTA = 0
                        xls_TAMANO_GRUPO = -1
                    End If
                    
                    
                    '--------
                    If (GRUPO_TEXTO = CStr(GRID.TextMatrix(Q_ROW1, Q_COL)) Or GRUPO_TEXTO = "xxxxxxxxx") And GRUPO_CUENTA = 0 Then
                    
                        objExcel.Cells(xls_ROW, xls_COL) = "'" & GRID.TextMatrix(Q_ROW1, Q_COL)
                        
    ''                    GRID.CellFontBold =
    ''                    GRID.CellForeColor = x_ForeColor
    ''                    GRID.CellBackColor = x_BackColor
                        '-----
                        '--COLOR AL TEXTO
        '                vp.TextColor = GRID.CellForeColor
                        '---ES_NEGRITA
                        objExcel.Cells(xls_ROW, xls_COL).Font.Bold = True 'GRID.CellFontBold
                        
                        objExcel.Cells(xls_ROW, xls_COL).HorizontalAlignment = COL_ALINEACION(CInt(Q_COL), True)
                        
                        If xls_TAMANO_GRUPO <> -1 Then
                            '--UNIR CELDAS
                            N_RANGO1 = objExcel.Cells(xls_ROW, xls_COL).Address
                            N_RANGO2 = objExcel.Cells(xls_ROW, xls_COL + (xls_TAMANO_GRUPO - Q_COL)).Address
                            N_RANGO_UNIR = N_RANGO1 & ":" & N_RANGO2
                            UNIR_CELDA objExcel, N_RANGO_UNIR
                        End If
                        
                        
                    End If
                    '----------------
                    If GRID.MergeCells = flexMergeFree Then GRUPO_CUENTA = 1
                    '--ancho de columna
                    If GRID.ColWidth(Q_COL) / 100 > 0.2 Then
                        objExcel.Columns(xls_COL).ColumnWidth = GRID.ColWidth(Q_COL) / 100
                    Else
                        objExcel.Columns(xls_COL).ColumnWidth = 0
                    End If
                    xls_COL = xls_COL + 1
                    '--
                End If
            Next Q_COL
            xls_ROW = xls_ROW + 1
            
            
            '---
        Next Q_ROW1
    Else
        '--COLOCANDO LOS ANCHOS DE LAS COLUMNAS
        For Q_ROW1 = 0 To 0
            xls_COL = xls_COL_INICIO
            For Q_COL = 1 To GRID.Cols - 1
                If GRID.ColWidth(Q_COL) <> 0 Then
                    objExcel.Columns(xls_COL).ColumnWidth = GRID.ColWidth(Q_COL) / 100
                    xls_COL = xls_COL + 1
                End If
            Next Q_COL
            '---
        Next Q_ROW1
    End If
        
End Sub



'------------------------------------------------------------------------------------
'''-----UNIR CELDAS
Private Sub UNIR_CELDA(objExcel As Object, pRango As String)
    With objExcel
        .Range(pRango).Select
        With .Selection
'            .HorizontalAlignment = -4108
'            .VerticalAlignment = -4107
            .WrapText = False
            .Orientation = 0
            .ShrinkToFit = False
            .MergeCells = False
        End With
        .Selection.Merge
    End With
End Sub

'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------

Private Sub PONER_ENCABEZADO_RST(objExcel As Object)
    'xCampos(?,0) Nombre Columna
    'xCampos(?,1) Campo Rst
    'xCampos(?,2) Alineacion 0:derecha; 1::centro; 2::izquierda
    'xCampos(?,3) Ancho
    
    Dim Q_ROW1&
    Dim xls_COL&, xls_ROW&
    Q_COL_INICIAL = 1
    xls_COL = xls_COL_INICIO
    xls_ROW = xls_ROW_INICIO - 2
    For Q_COL = 0 To UBound(ArrCampos)
        '--ancho
        If NulosN(ArrCampos(Q_COL, 3)) / 100 > 0.2 Then
            objExcel.Columns(xls_COL).ColumnWidth = NulosN(ArrCampos(Q_COL, 3)) / 100
        Else
            objExcel.Columns(xls_COL).ColumnWidth = 0
        End If
        
        objExcel.Cells(xls_ROW, xls_COL) = ArrCampos(Q_COL, 0) '--nombre
        objExcel.Cells(xls_ROW, xls_COL).Font.Bold = True
        'xlCenter= -4108
        'xlLeft= -4131
        'xlRight = -4152
        
        Select Case NulosN(ArrCampos(Q_COL, 2))
            Case 0: objExcel.Cells(xls_ROW, xls_COL).HorizontalAlignment = -4131 '--derecha
            Case 1: objExcel.Cells(xls_ROW, xls_COL).HorizontalAlignment = -4108 '--centro
            Case 2: objExcel.Cells(xls_ROW, xls_COL).HorizontalAlignment = -4152 '--izquierda
            Case Else
                objExcel.Cells(xls_ROW, xls_COL).HorizontalAlignment = -4131 '--derecha
        End Select
        xls_COL = xls_COL + 1
    Next Q_COL
    
End Sub



Private Sub PONER_DATOS_RST()
                
    On Error GoTo ERROR
    Dim N_RANGO1, N_RANGO2, N_RANGO_UNIR As String
    Dim xls_ROW&, xls_COL&, XCOL_ALINEACION&
    Dim nCol&
    
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")

    'determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1

    'abre el Libro
    objExcel.Workbooks.Add
    '------------------------------------
    Me.MousePointer = vbHourglass
    LBL(2).Visible = True
    SGI_JC.habilitar cmd, False
    BAND_INTERRUMPIR = False
    LBL(0).Caption = LBL(0).Tag:    LBL(1).Caption = LBL(1).Tag
    Me.Caption = LBL(0).Tag
    '------------------------------------
    With objExcel.ActiveSheet
        '-----------------------------------------------------------------
        PONER_EMPRESA objExcel
        '-----------------------------------------------------------------
        xls_ROW = xls_ROW_INICIO - 1
        xls_COL = xls_COL_INICIO
        
        '------
        PONER_ENCABEZADO_RST objExcel

        PgBar.Value = 0
        PgBar.Max = RstTmp.RecordCount
        RstTmp.MoveFirst
        Do While Not RstTmp.EOF
            DoEvents
            PgBar.Value = PgBar.Value + 1
            If BAND_INTERRUMPIR = True Then GoTo salir
            xls_COL = xls_COL_INICIO
            
            For nCol = 0 To UBound(ArrCampos)
                
                Select Case NulosN(ArrCampos(nCol, 2))
                    Case 0: XCOL_ALINEACION = -4131 '--derecha
                    Case 1: XCOL_ALINEACION = -4108 '--centro
                    Case 2: XCOL_ALINEACION = -4152 '--izquierda
                    Case Else
                        XCOL_ALINEACION = -4131 '--derecha
                End Select
                
                .Cells(xls_ROW, xls_COL).HorizontalAlignment = XCOL_ALINEACION
                
                If XCOL_ALINEACION = -4152 Then
                    .Cells(xls_ROW, xls_COL) = RstTmp.Fields(ArrCampos(nCol, 1))
                Else
                    .Cells(xls_ROW, xls_COL) = "'" & RstTmp.Fields(ArrCampos(nCol, 1))
                End If
               
                xls_COL = xls_COL + 1
                
            Next nCol
            
            xls_ROW = xls_ROW + 1
            RstTmp.MoveNext
        Loop
        '------
    End With
salir:
    
    If BAND_INTERRUMPIR = True Then
        MsgBox "El proceso de exportación se interrumpió", vbInformation, xTitulo
    Else
        MsgBox "El proceso de exportación terminó con exito", vbInformation, xTitulo
    End If
    objExcel.Visible = True
    objExcel.WindowState = 1
    Set objExcel = Nothing

    Me.MousePointer = vbDefault
    Unload Me
    Exit Sub
    
ERROR:
    BAND_INTERRUMPIR = False
    Me.MousePointer = vbDefault
    objExcel.Visible = True
    objExcel.WindowState = 1
    Set objExcel = Nothing
    SGI_JC.SHOW_ERROR "", "", True, IIf(Err.Number <> 50290, "", "No manipule el archivo hasta que termine de exportar!!!!")
    LBL(0).Caption = "Presione F9 para Salir..."
    Unload Me
End Sub

Private Sub PONER_EMPRESA(objExcel As Object)

    Dim N_RANGO1, N_RANGO2, N_RANGO_UNIR As String

    With objExcel.ActiveSheet
    '.Cells.Font.Name = "Arial"
        '.Cells.Font.Size = 8
        .Cells(1, 2) = NomEmp
        N_RANGO1 = .Cells(1, 2).Address:         N_RANGO2 = .Cells(1, 4).Address:
        N_RANGO_UNIR = N_RANGO1 & ":" & N_RANGO2
        UNIR_CELDA objExcel, N_RANGO_UNIR
        
        .Cells(2, 2) = "R.U.C. : " + NumRUC
        N_RANGO1 = .Cells(2, 2).Address:         N_RANGO2 = .Cells(2, 4).Address:
        N_RANGO_UNIR = N_RANGO1 & ":" & N_RANGO2
        UNIR_CELDA objExcel, N_RANGO_UNIR
        
        .Cells(3, 2) = Date
        N_RANGO1 = .Cells(3, 2).Address:         N_RANGO2 = .Cells(3, 4).Address:
        N_RANGO_UNIR = N_RANGO1 & ":" & N_RANGO2
        UNIR_CELDA objExcel, N_RANGO_UNIR
        
        .Cells(1, 2).HorizontalAlignment = -4131
        .Cells(2, 2).HorizontalAlignment = -4131
        .Cells(3, 2).HorizontalAlignment = -4131
        '-----------------------------------------------------------------
        '--DEL TITULO
        .Cells(5, 2) = T_TITULO:            .Cells(5, 2).Font.Bold = True
        .Cells(6, 2) = T_PERIODO:           .Cells(6, 2).Font.Bold = True
        .Cells(7, 2) = T_TITULO_1:          .Cells(7, 2).Font.Bold = True
        
        '-----------------------------------------------------------------
    End With
End Sub
