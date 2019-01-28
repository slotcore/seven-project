VERSION 5.00
Begin VB.Form FrmVSFlexGrid_Buscar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Datos"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cb 
      Height          =   315
      Left            =   975
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   45
      Width           =   3180
   End
   Begin VB.TextBox txt 
      Height          =   300
      Left            =   975
      TabIndex        =   0
      Text            =   "txt"
      Top             =   420
      Width           =   3180
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   2
      Top             =   870
      Width           =   1725
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Buscar Siguiente"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   870
      Width           =   1725
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   4035
      Y1              =   825
      Y2              =   825
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor: "
      Height          =   195
      Index           =   1
      Left            =   30
      TabIndex        =   5
      Top             =   540
      Width           =   450
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccionar: "
      Height          =   195
      Index           =   0
      Left            =   30
      TabIndex        =   4
      Top             =   195
      Width           =   930
   End
End
Attribute VB_Name = "FrmVSFlexGrid_Buscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''-----------
'----pBuscar DATOS POR VSFlexGrid
'----POR: JOHAN CASTRO
'----20/12/07


'----DESCRIPCION DE PARAMETROS A USAR EN FUNCION
'RECIBE_LINK_SEARCH(hWndFrmPadre , GRID_1 , X_ARRAY_TMP() , FILA_INICIO )
    'hWndFrmPadre   :: VALOR FORM_PADRE.hWnd QUE CONTENDRA AL FORMULARIO BUSCAR
    'GRID_1         :: OBJETO VSFlexGrid
    'X_ARRAY_TMP()  :: ARREGLO QUE CONTIENE (NOMBRE A BUSCAR, TIPO DE DATO, COLUMNA A BUSCAR)CADENA
    'FILA_INICIO    :: NUMWERO INDICA LA FILA ACTUAL SELECCIONADA DEL GRID
    'OBS:: TODOS LOS PARAMETROS SON OBLIGATORIOS

Dim SGI_JC As New SGI2_funciones.JC_Varios

'---------------------------------
Dim GRID As Object
Dim ARR_TMP() As String
Dim mRowInicio As Long '--INDICA LA ULTIMA POSICION DE LA FILA A BUSCAR
Dim mRowTemporal As Long '--INDICA LA ULTIMA FILA A BUSCAR
'-----------
 
Dim SeEjecuto As Boolean

Public Sub RECIBE_LINK_SEARCH(hWndFrmPadre As Long, GRID_1 As Object, X_ARRAY_TMP() As String, FILA_INICIO As Long)
    
    '---------------------------
    cb.Clear
    
    ReDim ARR_TMP(UBound(X_ARRAY_TMP()), 4)
    Dim POS_ARR As Integer
    Dim POS_DEFECTO As Integer
    
    mRowInicio = FILA_INICIO
    
    POS_ARR = 0
    POS_DEFECTO = -2
    
    '--0. CAMPO,
    '--1. COLUMNA, DEL GRID
    '--2. TIPO_DATO(C::CARACTER; N::NUMERICO, F::FECHA)
    '--3. PREDETERMINADO 0,-1 ESPECIFICA(EL CAMPO POR DEFECTO)
    For POS_ARR = 0 To UBound(X_ARRAY_TMP())
        ARR_TMP(POS_ARR, 0) = X_ARRAY_TMP(POS_ARR, 0)
        ARR_TMP(POS_ARR, 1) = X_ARRAY_TMP(POS_ARR, 1)
        ARR_TMP(POS_ARR, 2) = X_ARRAY_TMP(POS_ARR, 2)
        ARR_TMP(POS_ARR, 3) = X_ARRAY_TMP(POS_ARR, 3)
        cb.AddItem X_ARRAY_TMP(POS_ARR, 0)
        If X_ARRAY_TMP(POS_ARR, 3) = "-1" Then POS_DEFECTO = POS_ARR
        
'        POS_ARR = POS_ARR + 1
    Next
    
    If POS_DEFECTO <> -2 Then cb.ListIndex = POS_DEFECTO
    '---------------------------
    Set GRID = GRID_1
    
    mRowTemporal = -1
    
    SGI_JC.VentanaFlotante Me.hwnd, hWndFrmPadre
        
    DoEvents
End Sub

Private Sub cb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub cmd_Click(Index As Integer)
    Select Case Index
        Case 0
            If fValidarData() = False Then Exit Sub
            '----------
            Dim mRow As Long
            mRow = GRID.FindRow(Trim(txt.Text), mRowTemporal, ARR_TMP(cb.ListIndex, 1), False, False)
            If mRow = -1 Then
                MsgBox "No se encontró información alguna", vbInformation, "Buscar..."
                mRowTemporal = -1
                txt.Text = ""
            Else
                GRID.Row = mRow
                mRowTemporal = mRow + 1
            End If

            '-----------
        Case 1
            Set GRID = Nothing
            Unload Me
    End Select
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
    SeEjecuto = False
    
    txt.Text = ""

    SeEjecuto = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        '--interrumpir
        Unload Me
    ElseIf KeyCode = vbKeyF9 And Shift = 0 Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    
End Sub


Private Function fValidarData() As Boolean
    If cb.ListIndex = -1 Then
        MsgBox "Seleccione una Columna a Buscar", vbExclamation, "Buscar..."
        cb.SetFocus
        SendKeys "{F4}"
        Exit Function
    End If
    
    If Trim(txt.Text) = "" Then Exit Function
    
    Select Case UCase(ARR_TMP(cb.ListIndex, 2))
        Case "N"
            If IsNumeric(txt.Text) = False Then
                MsgBox "El Valos ingresado no es un número" + vbCr + "Ingrese un número correcto", vbExclamation, "Buscar..."
                Exit Function
            End If
        Case "C"
        
        Case "F"
            If IsDate(txt.Text) = False Then
                MsgBox "El Valos ingresado no es una Fecha" + vbCr + "Ingrese un numero Correcto", vbExclamation, "Buscar..."
                Exit Function
            End If
        Case Else
            
            Exit Function
            
        End Select
    
    fValidarData = True
    '---------
End Function

Private Sub pBuscar()
    On Error GoTo error
    Dim mRow As Long
    mRowInicio = -1
    mRow = GRID.FindRow(Trim(txt.Text), mRowInicio, ARR_TMP(cb.ListIndex, 1), False, False)
    If mRow = -1 Then
        MsgBox "No se encontró información alguna", vbInformation, "Buscar..."
        txt.Text = ""
    Else
        GRID.Row = mRow
        mRowTemporal = mRow + 1
    End If
    Exit Sub
error:
    MsgBox "Error: " & Err.Description & vbCr & "Origen: " & Err.Source & vbCr & "Número: " & Err.Number, vbCritical, "Error..."
    
End Sub

Private Sub txt_Change()
    If SeEjecuto = False Then Exit Sub
    If Trim(txt.Text) = "" Then
        mRowInicio = -1
        Exit Sub
    End If
    If fValidarData() = False Then Exit Sub
    pBuscar
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub
