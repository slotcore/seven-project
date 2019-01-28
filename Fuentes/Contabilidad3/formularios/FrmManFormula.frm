VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmManFormula 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración Estados Financieros - Editor de Fórmula"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "("
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   6150
      TabIndex        =   9
      Top             =   1185
      Width           =   390
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   6600
      TabIndex        =   8
      Top             =   1185
      Width           =   390
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      Caption         =   "Salir"
      Height          =   330
      Index           =   4
      Left            =   6135
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   420
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      Caption         =   "Aceptar"
      Height          =   330
      Index           =   3
      Left            =   6135
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   75
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      Caption         =   "Limpiar"
      Height          =   330
      Index           =   2
      Left            =   6150
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Limpiar Fórmula"
      Top             =   1545
      Width           =   855
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   1
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "FrmManFormula.frx":0000
      Top             =   3180
      Width           =   6990
   End
   Begin VB.TextBox txt 
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
      Height          =   1200
      Index           =   0
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "FrmManFormula.frx":0009
      Top             =   1920
      Width           =   6990
   End
   Begin VB.CommandButton cmd_oper 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   6600
      TabIndex        =   2
      ToolTipText     =   "Menos"
      Top             =   840
      Width           =   390
   End
   Begin VB.CommandButton cmd_oper 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   6150
      TabIndex        =   1
      ToolTipText     =   "Mas"
      Top             =   840
      Width           =   390
   End
   Begin VSFlex7Ctl.VSFlexGrid fg 
      Height          =   1830
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   5970
      _cx             =   10530
      _cy             =   3228
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmManFormula.frx":0012
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
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "Mas(+)"
      End
      Begin VB.Menu Menu1_4 
         Caption         =   "Menos(-)"
      End
   End
End
Attribute VB_Name = "FrmManFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_INDEX As Integer      '--INDICE DEL GRID
Dim F_OPERADOR As Boolean   '--INDICA PARA INGRESAR UN OPERADOR
Dim F_VALOR As Boolean      '--INDICA PARA INGRESAR UN VALOR
Dim M_ROW As Integer        '--INDICA LA FILA DEL GRID ORIGEN QUE TENDRA LA FORMULA
Dim M_COL As Integer        '--INDICA LA COLUMNA QUE CONTENDAR LA FORMULA

Dim N_FORMULA As String

Dim GRID As VSFlexGrid      '--INDICA LA GRILLA QUE CONTIENE LOS DATOS DE LA FORMULA



Public Sub RECIBE_LINK_FRM(GridTmp As VSFlexGrid, GridFormula As VSFlexGrid, _
                            Q_ROW As Long, Q_COL As Long, Formula As String, _
                            N_CAPTION As String, _
                            COL_ID As Integer, COL_VALOR As Integer, _
                            Optional SOLO_CTA As Boolean = False)

'CUENTA 2DIGITOS(SOLO ESTO SE CONSIDERA)
'SUBCUENTA 3 DIGITOS
'DIVISIONARIA VARIOS DIGITOS

'--GridTmp                  CONTIENE LOS REGISTROS A MOSTRAR EN EL GRID
'--GridFormula              CONTIENE EL REGISTRO DE LA FORMULA(SE USARA CUANDO SE GENERE LA FORMULA, SE ALMACENARA EN LA POSICION DE LA FORMULA SEGUN (Q_ROW,Q_COL))
'--Q_ROW, Q_COL,FORMULA     VALORES DE GridFormula
'--N_CAPTION                TEXTO A MOSTRAR EN EL TITULO DEL FORM.
'--COL_ID,COL_VALOR         COLUMNAS A MOSTRAR EN EL GRID, DEPENDE DE GridTmp

'--SOLO_CTA                 FALSE CARGAR TODOS LOS REGISTROS DE GridTmp
'--                         TRUE  CARGAR SOLO LAS CUENTAS Ej, 10,12,...
    Dim i_row As Long
    
    Set GRID = GridFormula
    
    M_ROW = Q_ROW
    M_COL = Q_COL
    M_INDEX = Index
    With GridTmp
        For i_row = 1 To .Rows - 1
            If .TextMatrix(i_row, COL_ID) <> "" And .TextMatrix(i_row, COL_VALOR) <> "" Then
                If SOLO_CTA = True Then
                    If Len(Trim(.TextMatrix(i_row, COL_ID))) <> 2 Then GoTo IR_SIG:
                End If
                fg.AddItem ""
                fg.TextMatrix(fg.Rows - 1, 1) = .TextMatrix(i_row, COL_ID)
                fg.TextMatrix(fg.Rows - 1, 2) = .TextMatrix(i_row, COL_VALOR)
            End If
IR_SIG:
        Next i_row
    End With
    
    LimpiaText txt, True
    
    If Formula <> "" Then
        N_FORMULA = Formula
        ARMAR_FORMULA
    End If
    
    Me.Caption = Me.Caption + " " + N_CAPTION
    
End Sub

Private Sub cmd_Click(Index As Integer)

    If Index = 0 Or Index = 1 Then
        txt(0).Text = txt(0).Text + " " + cmd(Index).Caption + " "
        txt(1).Text = txt(1).Text + cmd(Index).Caption
        Exit Sub
    End If
    
    Select Case Index
        Case 2 '--LIMPIAR
            If MsgBox("Seguro desea limpiar", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
            F_OPERADOR = False:         F_VALOR = False
            LimpiaText txt
        Case 3 '--ACEPTAR
        
            If F_OPERADOR = True Then
                MsgBox "No puede continuar, a ingresado un operador" + _
                vbCr + "Ingrese un Valor, luego proceda... ", vbExclamation, xTitulo
                
                fg.SetFocus
                Exit Sub
            End If
        
            If MsgBox("Seguro desea agregar la Fórmula", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
            GRID.TextMatrix(M_ROW, M_COL) = txt(1).Text
            Unload Me
        Case 4 '--SALIR
            Unload Me
    End Select

End Sub

Private Sub cmd_oper_Click(Index As Integer)
    If F_OPERADOR = True Then
        MsgBox "No puede ingresar otro Operador" + _
        vbCr + "Ingrese un Valor, luego proceda... ", vbExclamation, xTitulo
        
        fg.SetFocus
        Exit Sub
    End If
    If txt(0).Text = "" Then
        MsgBox "Ingrese primero un Valor" + _
        vbCr + "Ingrese un Valor, luego proceda... ", vbExclamation, xTitulo
        Exit Sub
    End If
    
    F_VALOR = False
    F_OPERADOR = True

    txt(0).Text = txt(0).Text + " " + cmd_oper(Index).Caption + " "
    txt(1).Text = txt(1).Text + cmd_oper(Index).Caption
        
End Sub

Private Sub fg_DblClick()
    If fg.Row < 0 Then Exit Sub
    If F_VALOR = True Then
        MsgBox "No puede ingresar otro Valor" + _
        vbCr + "Ingrese un Operador, luego proceda... ", vbExclamation, xTitulo
        Exit Sub
    End If
    txt(0).Text = txt(0).Text + fg.TextMatrix(fg.Row, 2) + " "
    txt(1).Text = txt(1).Text + fg.TextMatrix(fg.Row, 1)
    F_OPERADOR = False
    F_VALOR = True
End Sub

Private Sub Fg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu Menu1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
'    If KeyCode = 107 Then cmd_Click 0
'    If KeyCode = 109 Then cmd_Click 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_Load()
    CentrarFrm Me
    habilitar_Locked txt, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    F_OPERADOR = False:         F_VALOR = False
    Set GRID = Nothing
End Sub

Private Sub Menu1_1_Click()
    fg_DblClick
End Sub

Private Sub Menu1_3_Click()
    cmd_oper_Click 0
End Sub

Private Sub Menu1_4_Click()
    cmd_oper_Click 1
End Sub


Private Sub ARMAR_FORMULA()
    '--ESTA FUNCION CREARA LA FORMULA EN FUNCION A LA CADENA DE ID'S
    '--  CADENA DE ID'S     =>>>>> FORMULA
    '--16+17+18+19          =>>>>> Contigencias + Intereses Minoritarios + PATRIMONIO NETO + PATRIMONIO NETO
    '-------------

    Dim Q_POS As Integer
    Dim N_ID As String
    Dim N_VALOR As String
    
    Do While N_FORMULA <> ""
        DoEvents
INICIAR_OTRA_VEZ:
''''        '--DEL OPERADOR ( ó )
''''
''''        If Mid(N_FORMULA, 1, 1) = "(" Or Mid(N_FORMULA, 1, 1) = ")" Then
''''                N_ID = Mid(N_FORMULA, 1, 1)
''''                txt(0).Text = txt(0).Text + N_VALOR + " + "
''''                txt(1).Text = txt(1).Text + N_ID + "+"
''''
''''                N_FORMULA = Right(N_FORMULA, Len(N_FORMULA) - 1)
''''                GoTo INICIAR_OTRA_VEZ
''''        End If


        '--DEL OPERADOR (+)
        Q_POS = InStr(N_FORMULA, "+")
        If Q_POS <> 0 Then
            N_ID = Mid(N_FORMULA, 1, Q_POS - 1)
            If InStr(N_ID, "-") <> 0 Then
                '--N_FORMULA = 22-12-11+17-18+19
                '--N_ID =22-12-11
                GoTo INICIAR_MENOS
            End If
            N_VALOR = GRID_BUSCAR_VALOR(fg, 1, N_ID, False, 2)
            If N_VALOR <> "-1" Then
                txt(0).Text = txt(0).Text + N_VALOR + " + "
                txt(1).Text = txt(1).Text + N_ID + "+"
                F_OPERADOR = True
            End If
            
            N_FORMULA = Right(N_FORMULA, Len(N_FORMULA) - Q_POS)
            If N_FORMULA <> "" Then GoTo INICIAR_OTRA_VEZ
            
        End If
        '--RENOMBRAR FORMULA
        'DEL OPERADOR (-)
INICIAR_MENOS:
        Q_POS = InStr(N_FORMULA, "-")
        If Q_POS <> 0 Then
            N_ID = Mid(N_FORMULA, 1, Q_POS - 1)
            N_VALOR = GRID_BUSCAR_VALOR(fg, 1, N_ID, False, 2)
            
            If N_VALOR <> "-1" Then
                txt(0).Text = txt(0).Text + N_VALOR + " - "
                txt(1).Text = txt(1).Text + N_ID + "-"
                F_OPERADOR = True
            End If
            
            N_FORMULA = Right(N_FORMULA, Len(N_FORMULA) - Q_POS)
            
            If N_FORMULA <> "" Then GoTo INICIAR_OTRA_VEZ
            
        End If
    
        If IsNumeric(N_FORMULA) = True Then
        
            N_VALOR = GRID_BUSCAR_VALOR(fg, 1, N_FORMULA, False, 2)
            
            If N_VALOR <> "-1" Then
                txt(0).Text = txt(0).Text + N_VALOR
                txt(1).Text = txt(1).Text + N_FORMULA
                F_OPERADOR = False
                F_VALOR = True
            End If
            N_FORMULA = ""
        Else
            Exit Do
        End If
    
        
        
    Loop

End Sub

