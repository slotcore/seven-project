VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Begin VB.Form FrmManBalanceFormula 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Fórmula"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Salir"
      Height          =   330
      Index           =   4
      Left            =   6135
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   450
      Width           =   855
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      BackColor       =   &H00C0C0C0&
      Caption         =   "Limpiar"
      Height          =   330
      Index           =   2
      Left            =   6135
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1455
      Width           =   855
   End
   Begin VB.TextBox txt 
      Height          =   360
      Index           =   1
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "FrmManBalanceFormula.frx":0000
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
      Text            =   "FrmManBalanceFormula.frx":0009
      Top             =   1920
      Width           =   6990
   End
   Begin VB.CommandButton cmd 
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
      Left            =   6615
      TabIndex        =   2
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmd 
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
      Left            =   6135
      TabIndex        =   1
      Top             =   1080
      Width           =   375
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmManBalanceFormula.frx":0012
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
Attribute VB_Name = "FrmManBalanceFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim M_INDEX As Integer '--INDICE DEL GRID
Dim F_OPERADOR As Boolean '--INDICA PARA INGRESAR UN OPERADOR
Dim F_VALOR As Boolean '--INDICA PARA INGRESAR UN VALOR
Dim M_ROW As Integer '--INDICA LA FILA DEL GRID ORIGEN QUE TENDRA LA FORMULA

Dim N_FORMULA As String


Public Sub RECIBE_LINK_FRM(index As Integer, X_ROW As Integer)
    M_ROW = X_ROW
    M_INDEX = index
    With FrmManBalance.fg(index)
        For X_ROW = 1 To .Rows - 1
            If .TextMatrix(X_ROW, 3) <> "" And .TextMatrix(X_ROW, 4) <> "" Then
                fg.AddItem ""
                fg.TextMatrix(fg.Rows - 1, 1) = .TextMatrix(X_ROW, 3)
                fg.TextMatrix(fg.Rows - 1, 2) = .TextMatrix(X_ROW, 4)
            End If
        Next X_ROW
    End With
    LimpiaText txt, True
    
    If FrmManBalance.fg(index).TextMatrix(M_ROW, 7) <> "" Then
        N_FORMULA = FrmManBalance.fg(index).TextMatrix(M_ROW, 7)
        ARMAR_FORMULA
    End If
    
    Me.Caption = "Editor de Fórmula - " + IIf(M_INDEX = 0, " Activo", "Pasivo")
    
End Sub

Private Sub cmd_Click(index As Integer)

If index = 0 Or index = 1 Then
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
End If

Select Case index
    Case 0
        txt(0).Text = txt(0).Text + " + "
        txt(1).Text = txt(1).Text + "+"
    Case 1
        txt(0).Text = txt(0).Text + " - "
        txt(1).Text = txt(1).Text + "-"
    Case 2
        If MsgBox("Seguro desea limpiar", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
        F_OPERADOR = False:         F_VALOR = False
        LimpiaText txt
    Case 3
    
        If F_OPERADOR = True Then
            MsgBox "No puede continuar, a ingresado un operador" + _
            vbCr + "Ingrese un Valor, luego proceda... ", vbExclamation, xTitulo
            
            fg.SetFocus
            Exit Sub
        End If
    
        If MsgBox("Seguro desea agregar la Fórmula", vbQuestion + vbYesNo, xTitulo) = vbNo Then Exit Sub
        FrmManBalance.fg(M_INDEX).TextMatrix(M_ROW, 7) = txt(1).Text
        Unload Me
    Case 4
        Unload Me
End Select

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

Private Sub Fg_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
End Sub

Private Sub Menu1_1_Click()
    fg_DblClick
End Sub

Private Sub Menu1_3_Click()
    cmd_Click 0
End Sub

Private Sub Menu1_4_Click()
    cmd_Click 1
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
INICIAR_OTRA_VEZ:
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
        End If
    
        If IsNumeric(N_FORMULA) = True Then
        
            N_VALOR = GRID_BUSCAR_VALOR(fg, 1, N_ID, False, 2)
            
            If N_VALOR <> "-1" Then
                txt(0).Text = txt(0).Text + N_VALOR
                txt(1).Text = txt(1).Text + N_FORMULA
                F_OPERADOR = False
                F_VALOR = True
            End If
            N_FORMULA = ""
        End If
    
    Loop

End Sub
