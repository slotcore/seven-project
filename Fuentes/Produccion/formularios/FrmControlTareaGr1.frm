VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmControlTareaGr1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupo"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   50
      TabIndex        =   17
      Top             =   -50
      Width           =   6345
      Begin VB.CheckBox chkAuto 
         Caption         =   "Cálculo A&utomatico"
         Height          =   195
         Left            =   4590
         TabIndex        =   21
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Deseleccionar Todos"
         Height          =   225
         Left            =   2010
         TabIndex        =   19
         Top             =   200
         Width           =   1965
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Seleccionar &Todos"
         Height          =   225
         Left            =   60
         TabIndex        =   18
         Top             =   200
         Width           =   1785
      End
   End
   Begin VB.Frame FraEditor 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   45
      TabIndex        =   5
      Top             =   3015
      Visible         =   0   'False
      Width           =   5205
      Begin VB.CommandButton CmdEditor 
         Caption         =   "&Salir"
         Height          =   420
         Index           =   2
         Left            =   4200
         TabIndex        =   9
         Top             =   2415
         Width           =   945
      End
      Begin VB.CommandButton CmdEditor 
         Caption         =   "Eliminar"
         Height          =   420
         Index           =   1
         Left            =   4200
         TabIndex        =   8
         Top             =   1815
         Width           =   945
      End
      Begin VB.CommandButton CmdEditor 
         Caption         =   "Agregar"
         Height          =   420
         Index           =   0
         Left            =   4200
         TabIndex        =   7
         Top             =   1320
         Width           =   945
      End
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   4935
         Picture         =   "FrmControlTareaGr1.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   6
         ToolTipText     =   "Cerrar"
         Top             =   30
         Width           =   195
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   2445
         Index           =   1
         Left            =   60
         TabIndex        =   10
         Top             =   360
         Width           =   4050
         _cx             =   7144
         _cy             =   4313
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   12
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmControlTareaGr1.frx":02EC
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
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblTotal(3)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   4185
         TabIndex        =   15
         Top             =   975
         Width           =   870
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total P.N"
         Height          =   195
         Index           =   2
         Left            =   4185
         TabIndex        =   14
         Top             =   780
         Width           =   675
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblTotal(1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   4185
         TabIndex        =   13
         Top             =   540
         Width           =   870
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total P.B"
         Height          =   195
         Index           =   0
         Left            =   4185
         TabIndex        =   12
         Top             =   360
         Width           =   660
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         Index           =   2
         X1              =   4140
         X2              =   4140
         Y1              =   345
         Y2              =   2700
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   15
         Y1              =   15
         Y2              =   5070
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   0
         X1              =   30
         X2              =   5700
         Y1              =   2865
         Y2              =   2880
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -150
         X2              =   5895
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5190
         X2              =   5190
         Y1              =   -120
         Y2              =   4770
      End
      Begin VB.Label LblTituloFrame 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Registros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   75
         TabIndex        =   11
         Top             =   60
         Width           =   810
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   270
         Index           =   1
         Left            =   30
         Top             =   15
         Width           =   5130
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid fg 
      Height          =   2010
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   465
      Width           =   6360
      _cx             =   11218
      _cy             =   3545
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmControlTareaGr1.frx":03D6
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
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   45
      TabIndex        =   1
      Top             =   2445
      Width           =   6345
      Begin VB.CommandButton Cmd 
         Caption         =   "&Seleccionar"
         Height          =   330
         Index           =   2
         Left            =   1140
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Eliminar Personal"
         Top             =   135
         Width           =   1065
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "&Agregar"
         Height          =   330
         Index           =   0
         Left            =   60
         TabIndex        =   2
         ToolTipText     =   "Agregar Personal"
         Top             =   135
         Width           =   1065
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "Eliminar Todos"
         Height          =   330
         Index           =   3
         Left            =   3435
         TabIndex        =   20
         ToolTipText     =   "Agregar Personal"
         Top             =   135
         Width           =   1200
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "Elimi&nar"
         Height          =   330
         Index           =   1
         Left            =   2400
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Eliminar Personal"
         Top             =   135
         Width           =   1035
      End
      Begin VB.Label lblTotalGr 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   195
         Left            =   4710
         TabIndex        =   4
         Top             =   180
         Width           =   315
      End
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
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "FrmControlTareaGr1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMCONTROLTAREAGR.FRM
'* Tipo             : FORMULARIO
'* Descripcion      :
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 31/10/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim Agregando As Boolean
Dim QueHace As Integer
'--fila que origina la muestra de datos, viene de FrmControlTarea.fg1
Dim mCodigoPrincipal
Dim RstGrDetalle As New ADODB.Recordset
Dim RstGrDetalleTara As New ADODB.Recordset
Dim sCantidad As Double
Dim mIdUniMed As Long                   ' codigo de la unidad de medida
Dim mRowAdd As Double                   ' identificador unico por fila cuando se agrege una tarea
Dim fCalculoAuto As Boolean             ' establecera el calculo en automatico
Public fDesactivarAuto As Boolean

'*****************************************************************************************************
'* Nombre           : pRecibeLink
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : Iniciar la Carga del Formulario del grupo, colocar este formulario como flotante
'*                    Cargar los registros del grupo en el grid
'* Paranetros       : NOMBRE           |  TIPO     |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    hWndFrmPadre     |  Long     |  origen del formulario .hWnd
'*                    mRowFila         |           |  fila seleccionada de ingreso tareas
'*                    fCalculoCantidad |  Boolean  |  indicara el tipo de calculo. true::automatico;
'*                                                    false::manual
'*                    fCalculoHora     |  Boolean  |
'* Devuelve         :
'*****************************************************************************************************
Public Sub pRecibeLink(hWndFrmPadre As Long, mRowFila, fCalculoCantidad As Boolean, fCalculoHora As Boolean)
    ' del tipo de calculo(distribuir la cantidad por grupo)
    fCalculoAuto = fCalculoCantidad
    
    ' mRowFila codigo del FlexGrid
    sCantidad = NulosN(FrmControlTarea1.Fg1.TextMatrix(FrmControlTarea1.Fg1.Row, 8))
    
    ' unidad de medida
    mIdUniMed = NulosN(FrmControlTarea1.Fg1.TextMatrix(FrmControlTarea1.Fg1.Row, 15))
    mCodigoPrincipal = mRowFila
    
    ' poner el titulo del formulario
    Me.Caption = FrmControlTarea1.Fg1.TextMatrix(FrmControlTarea1.Fg1.Row, 3) & "    Cant: " & Format(sCantidad, FORMAT_MONTO) & " " & FrmControlTarea1.Fg1.TextMatrix(FrmControlTarea1.Fg1.Row, 9)
    If NulosC(Me.Caption) = "" Then Me.Caption = "Grupo"
    
    ' traspasar el recorsed a uno temporal para efectos del calculo
    Set RstGrDetalle = FrmControlTarea1.RstGrDet
    Set RstGrDetalleTara = FrmControlTarea1.RstGrDetTara
    
    ' filtrar el recordset segun fila seleccionada
    RstGrDetalle.Filter = "codigo = " & mCodigoPrincipal
    RstGrDetalleTara.Filter = "codigo = " & mCodigoPrincipal & " and tipo=1 "
    
    ' ver si se puede habilitar los botones
    If FrmControlTarea1.Cmd(0).Enabled = False Then
        If FrmControlTarea1.FraEditor.Visible = False Then
            QueHace = 3
            habilitar Cmd, False
            fg(0).SelectionMode = flexSelectionByRow
        End If
    Else
        QueHace = 2         ' modificar
        habilitar Cmd, True
        fg(0).SelectionMode = flexSelectionFree
    End If
    
    ' activar por defecto (deshabilitar el calculo automatico)
    chkAuto.Value = 0
    If fCalculoHora = True Then
        If RstGrDetalle.RecordCount > 0 Then
            RstGrDetalle.MoveFirst
            Do While Not RstGrDetalle.EOF
                If NulosN(RstGrDetalle("activo")) <> 0 Then
                    If IsDate(FrmControlTarea1.Fg1.TextMatrix(FrmControlTarea1.Fg1.Row, 6)) = True Then '--hora inicio
                        RstGrDetalle("horini") = CDate(FrmControlTarea1.Fg1.TextMatrix(FrmControlTarea1.Fg1.Row, 6))
                    Else
                        RstGrDetalle("horini") = Null
                    End If
                    If IsDate(FrmControlTarea1.Fg1.TextMatrix(FrmControlTarea1.Fg1.Row, 7)) = True Then '--hora final
                        RstGrDetalle("horfin") = CDate(FrmControlTarea1.Fg1.TextMatrix(FrmControlTarea1.Fg1.Row, 7))
                    Else
                        RstGrDetalle("horfin") = Null
                    End If
                Else
                    RstGrDetalle("horini") = Null
                    RstGrDetalle("horfin") = Null
                End If
                RstGrDetalle.MoveNext
            Loop
        End If
        pCargarDatos
    End If
    
    If fCalculoAuto = True Then
        pRecalcular
    Else
        pCargarDatos
    End If
    
    ' posicionar la ventana sobre la principal
    VentanaFlotante Me.hWnd, hWndFrmPadre
    
    ' posicionar la ventana segun la cantidad de registros de tareas
    Me.Height = 3465
    If FrmControlTarea1.Fg1.Row <= 9 Then
        Me.Top = 3700 - FrmControlTarea1.Top  '1550
        Me.Left = 5250 '2500 '
    Else
        Me.Top = -100 - FrmControlTarea1.Top '-2150
        Me.Left = 5250 '
    End If
    fCalculoCantidad = False
    fCalculoHora = False
End Sub

Private Sub cmd_Click(Index As Integer)
    If QueHace = 3 Then Exit Sub
    Select Case Index
        Case 0 ' agregar
            pRegistroAdd 0
        Case 1 ' eliminar
            pRegistroDel 0
        Case 2 ' Agregar por Lista
            listarEmpleados
        Case 3 ' Eliminar todos
            Dim A As Integer
            Dim num As Integer
            
            num = fg(0).Rows
            For A = 1 To num - 1
                fg(0).Select 1, 1, 1, fg(0).Cols - 1
                pRegistroDel 0
            Next A
    End Select
End Sub

Private Sub Fg_EnterCell(Index As Integer)
    If QueHace = 3 Then
        fg(Index).Editable = flexEDNone
        Exit Sub
    End If
    If Index = 0 Then
        If fg(0).Col = 5 Then
            If NulosN(fg(0).TextMatrix(fg(0).Row, 1)) = 0 Then
                fg(Index).Editable = flexEDNone
            Else
                fg(Index).Editable = flexEDKbdMouse
            End If
        Else
            fg(Index).Editable = flexEDKbdMouse
        End If
    Else
        fg(Index).Editable = flexEDKbdMouse
    End If
End Sub

Private Sub fg_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If QueHace = 3 Then Exit Sub
    If Index = 0 Then
        If Col <> 5 And Col <> 9 And Col <> 10 Then Exit Sub
        
        If Col = 9 Or Col = 10 Then
            If fg(0).TextMatrix(Row, Col) = "  :  " Then
                fg(0).TextMatrix(Row, Col) = ""
            ElseIf IsDate(fg(0).TextMatrix(Row, Col)) = False Then
                fg(0).TextMatrix(Row, Col) = ""
            Else
                fg(0).TextMatrix(Row, Col) = Format(fg(0).TextMatrix(Row, Col), FORMAT_HORA_SIN_SEGUNDO)
            End If
            
        Else
            If IsNumeric(fg(0).TextMatrix(Row, Col)) = False Then
                MsgBox "El número es incorrecto", vbExclamation, xTitulo
                Exit Sub
            End If
            
        End If
        
        ' aplicando filtro
        RstGrDetalle.Filter = "codigo = " & mCodigoPrincipal
        If RstGrDetalle.RecordCount <> 0 Then RstGrDetalle.MoveFirst
        RstGrDetalle.Find "idemp= " & NulosN(fg(0).TextMatrix(Row, 4))
        If RstGrDetalle.EOF = False And RstGrDetalle.BOF = False Then
            RstGrDetalle("cantbrut") = NulosN(fg(0).TextMatrix(Row, 3))
            RstGrDetalle("cant") = NulosN(fg(0).TextMatrix(Row, 5))
            ' colocando las horas
            If IsDate(fg(0).TextMatrix(Row, 9)) = True Then RstGrDetalle("horini") = CDate(fg(0).TextMatrix(Row, 9))
            If IsDate(fg(0).TextMatrix(Row, 10)) = True Then RstGrDetalle("horfin") = CDate(fg(0).TextMatrix(Row, 10))
            If IsDate(fg(0).TextMatrix(Row, 9)) = False Then RstGrDetalle("horini") = Null
            If IsDate(fg(0).TextMatrix(Row, 10)) = False Then RstGrDetalle("horfin") = Null
        End If
        FrmControlTarea1.Fg1.TextMatrix(FrmControlTarea1.Fg1.Row, 8) = GRID_SUMAR_COL(fg(0), 5)
    ElseIf Index = 1 Then
        ' aplicando filtro
        RstGrDetalleTara.Filter = "codigo = " & mCodigoPrincipal & " and tipo=1 and idemp = " & NulosN(fg(0).TextMatrix(fg(0).Row, 4))
        If RstGrDetalleTara.RecordCount <> 0 Then RstGrDetalleTara.MoveFirst
        RstGrDetalleTara.Find "item = " & NulosN(fg(1).TextMatrix(Row, 7))
        If RstGrDetalleTara.EOF = False And RstGrDetalleTara.BOF = False Then
            RstGrDetalleTara("cantidad") = NulosN(fg(1).TextMatrix(Row, 2))
            RstGrDetalleTara("pesouni") = NulosN(fg(1).TextMatrix(Row, 4))
            RstGrDetalleTara("pesotara") = NulosN(RstGrDetalleTara("pesouni")) * NulosN(RstGrDetalleTara("cantidad"))
            RstGrDetalleTara("pesobrut") = NulosN(fg(1).TextMatrix(Row, 1))
            RstGrDetalleTara("pesonet") = NulosN(RstGrDetalleTara("pesobrut")) - NulosN(RstGrDetalleTara("pesotara"))
            fg(1).TextMatrix(Row, 5) = NulosN(RstGrDetalleTara("pesonet"))
        End If
        If IsNumeric(fg(1).TextMatrix(Row, Col)) = False Then fg(1).TextMatrix(Row, Col) = 0
        lblTotal(1).Caption = Format(GRID_SUMAR_COL(fg(1), 1), FORMAT_MONTO)
        lblTotal(3).Caption = Format(GRID_SUMAR_COL(fg(1), 5), FORMAT_MONTO)
    End If
End Sub

Private Sub Fg_Click(Index As Integer)
    If Index = 1 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    If fg(0).Col <> 1 Then Exit Sub
    
    ' ir al primer registro
    If RstGrDetalle.RecordCount = 0 Then Exit Sub
    RstGrDetalle.MoveFirst
    ' buscar el personal para actualizar si esta activado o no
    RstGrDetalle.Find "idemp=" & NulosN(fg(0).TextMatrix(fg(0).Row, 4))
    If RstGrDetalle.EOF = False And RstGrDetalle.BOF = False Then
        RstGrDetalle("activo") = fg(0).TextMatrix(fg(0).Row, 1)
        ' refrescar los calculos
        If chkAuto.Value = 1 Then
            pRecalcular
        Else
            fg(0).TextMatrix(fg(0).Row, 3) = 0
            fg(0).TextMatrix(fg(0).Row, 5) = 0
            RstGrDetalle("cant") = 0
            RstGrDetalle("cantbrut") = 0
            ' eliminando el registro de taras
            RstGrDetalleTara.Filter = "codigo = " & mCodigoPrincipal & " and tipo=1 and idemp=" & NulosN(fg(0).TextMatrix(fg(0).Row, 4))
            If RstGrDetalleTara.RecordCount <> 0 Then RstGrDetalleTara.MoveFirst
            Do While Not RstGrDetalleTara.EOF
                RstGrDetalleTara.Delete
                RstGrDetalleTara.MoveNext
            Loop
        End If
    End If
End Sub

Private Sub listarEmpleados()
    '===================================================================================================
    'Creado : 16/04/11 Por: Jose Chacon
    'Propósito: Mostras listado de personal para seleccionar
    '
    'Entradas:  Ninguno
    '
    'Resultados: Lista de empleados en pantalla
    '
    '===================================================================================================

    If QueHace = 3 Then Exit Sub
    Dim nSQL As String
    Dim nSQLId As String
    Dim nSQLTmp  As String
    Dim nTitulo As String
    Dim xform As New eps_librerias.FormSeleccion
    Dim xRs As New ADODB.Recordset
    Dim xCampos(5, 4) As String
    
    xCampos(0, 0) = "Código":               xCampos(0, 1) = "codemp":       xCampos(0, 2) = "800":      xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
    xCampos(1, 0) = "Apellidos y Nombres":  xCampos(1, 1) = "nombres":      xCampos(1, 2) = "4000":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
    xCampos(2, 0) = "DNI":                  xCampos(2, 1) = "numdoc":       xCampos(2, 2) = "1000":     xCampos(2, 3) = "C":    xCampos(1, 4) = "C"
    xCampos(3, 0) = "Fch. Ingeso":          xCampos(3, 1) = "fching":       xCampos(3, 2) = "1000":     xCampos(3, 3) = "C":    xCampos(1, 4) = "C"
    xCampos(4, 0) = "Id":                   xCampos(4, 1) = "idemp":        xCampos(4, 2) = "900":      xCampos(4, 3) = "C":    xCampos(2, 4) = "C"
    
    If fg(0).Rows = fg(0).FixedRows Then fg(0).Rows = fg(0).Rows + 1
    ' generar la lista de personal para no considerar en la lista
    nSQLId = GRID_GENERAR_SQL_ID(fg(0), 4, " AND pla_empleados.id", "NOT IN", True)
    If NulosC(fg(0).TextMatrix(fg(0).Row, fg(0).Col)) <> "" Then
        'nSQLTmp = " AND UCASE([pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom]) LIKE '%" & UCase(NulosC(fg(0).TextMatrix(fg(0).Row, fg(0).Col))) & "%'"
    End If
    
    ' generar la consulta
    'SELECT " & mCodigoPrincipal & " as codigo, pla_empleados.id AS idemp, 0 AS idgrupo, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, 0 AS cant,0 as cantbrut, -1 AS activo, pla_empleados.codigo as codemp
    nSQL = "SELECT 0 AS xsel, " & mCodigoPrincipal & " as codigo, pla_empleados.id AS idemp, [pla_empleados].[nombre] AS nombres, pla_empleados.codigo as codemp,format(pla_empleados.fching,'dd/mm/yyyy') as fching,pla_empleados.numdoc " _
        + vbCr + " FROM (pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper " _
        + vbCr + " WHERE pla_empleados.fchcese is null and (((pro_empdet.idfun) = 6)) " & nSQLId & nSQLTmp _
        + vbCr + " ORDER BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom];"
    nTitulo = "Buscando Personal"

    xform.SQLCad = nSQL
        
    xform.titulo = "Buscando Personal"
    Set xform.Coneccion = xCon
    Set xRs = Nothing
    Set xRs = xform.seleccionar(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            Dim A As Integer
            'fg(0).Rows = 1
            For A = 1 To xRs.RecordCount
                ' agregando los datos al rst temporal
                RstGrDetalle.AddNew
                RstGrDetalle("codigo") = NulosC(xRs("codigo"))
                RstGrDetalle("idemp") = xRs("idemp")
                RstGrDetalle("idgrupo") = 0
                RstGrDetalle("nombres") = NulosC(xRs("nombres"))
                RstGrDetalle("cant") = 0
                RstGrDetalle("cantbrut") = 0
                RstGrDetalle("activo") = -1
                
                xRs.MoveNext
                If xRs.EOF = True Then Exit For
            Next A
            
            If chkAuto.Value = 1 Then
                pRecalcular
            Else
                pCargarDatos
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub seleccionartodos(seleccionar As Boolean)
    Dim A As Integer
    Dim valor As Double
    
    If seleccionar Then valor = -1 Else valor = 0
    
    If RstGrDetalle.RecordCount <> 0 Then
        ' agregando los datos al rst temporal
        RstGrDetalle.MoveFirst
        For A = 1 To fg(0).Rows - 1
            RstGrDetalle("activo") = valor
            RstGrDetalle.MoveNext
            If RstGrDetalle.EOF = True Then Exit For
        Next A
    End If
    pCargarDatos
End Sub

Private Sub iniciarCampos()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE EJECUTE EL FORMULARIO
    mRowAdd = -9999
    fg(0).ColWidth(3) = 0   ' peso bruto
    fg(0).ColWidth(4) = 0   ' idpersonal
    fg(0).ColWidth(6) = 0   ' lote
    fg(0).ColWidth(7) = 0   ' tarea
    fg(0).ColWidth(8) = 0   ' producto
    fg(0).ColWidth(11) = 0  ' idrec
    fg(0).ColWidth(12) = 0  ' idtarea
    fg(0).ColWidth(13) = 0  ' codigo
    fg(1).ColWidth(6) = 0   ' id tara
    fg(1).ColWidth(7) = 0   ' codigo
    
    GRID_COMBOLIST fg(0), 2 ' personal
'    GRID_COMBOLIST fg(0), 5 ' peso neto
    GRID_COMBOLIST fg(0), 7 ' producto
    GRID_COMBOLIST fg(0), 8 ' tarea
    
    fg(0).ColEditMask(9) = "##:##"  ' h.inicio
    fg(0).ColEditMask(10) = "##:##" ' h.fin
    
    GRID_COMBOLIST fg(1), 3         ' peso - tara
    
    fg(0).ColFormat(3) = FORMAT_MONTO
    fg(0).ColFormat(5) = FORMAT_MONTO
    
    fg(1).ColFormat(1) = FORMAT_MONTO ' peso bruto
    fg(1).ColFormat(4) = FORMAT_MONTO ' peso unit
    fg(1).ColFormat(5) = FORMAT_MONTO ' peso neto
    
    ' limpiando los totales
    lblTotal(1).Caption = ""
    lblTotal(3).Caption = ""
    
    QueHace = FrmControlTarea1.QueHace

    If QueHace = 3 Then
        fg(0).AllowUserResizing = flexResizeColumns
        fg(0).AutoSearch = flexSearchFromTop
        fg(0).ExplorerBar = flexExSortShowAndMove
        fg(0).SelectionMode = flexSelectionByRow
        fg(0).ForeColorSel = &H80000005
        fg(0).BackColorSel = &H80&
    Else
        fg(0).AutoSearch = flexSearchNone
    End If
End Sub

Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Col <> 2 And Col <> 3 And Col <> 5 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    Dim xRs As New ADODB.Recordset
    Dim nSQL As String
    Dim nSQLId As String
    Dim nSQLTmp  As String
    Dim nTitulo As String
    ' del personal
    If Index = 0 Then
        Select Case Col
            Case 2 ' personal
                ReDim xCampos(3, 4) As String
                xCampos(0, 0) = "Cod. Empleado":        xCampos(0, 1) = "codemp":       xCampos(0, 2) = "1500":     xCampos(0, 3) = "N":    xCampos(0, 4) = "N"
                xCampos(1, 0) = "Apellidos y Nombres":  xCampos(1, 1) = "nombres":      xCampos(1, 2) = "4500":     xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
                xCampos(2, 0) = "Id":                   xCampos(2, 1) = "idemp":        xCampos(2, 2) = "1000":     xCampos(2, 3) = "N":    xCampos(2, 4) = "N"
                
                ' generar la lista de personal para no considerar en la lista
                nSQLId = GRID_GENERAR_SQL_ID(fg(0), 4, " AND pla_empleados.id", "NOT IN", True)
                If NulosC(fg(0).TextMatrix(Row, Col)) <> "" Then
                    nSQLTmp = " AND UCASE([pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom]) LIKE '%" & UCase(NulosC(fg(0).TextMatrix(Row, Col))) & "%'"
                End If
                
                ' generar la consulta
                nSQL = "SELECT " & mCodigoPrincipal & " as codigo, pla_empleados.id AS idemp, 0 AS idgrupo, [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom] AS nombres, 0 AS cant,0 as cantbrut, -1 AS activo, pla_empleados.codigo as codemp " _
                    + vbCr + " FROM (pla_empleados INNER JOIN pro_emp ON pla_empleados.id = pro_emp.idemp) INNER JOIN pro_empdet ON pro_emp.id = pro_empdet.idper " _
                    + vbCr + " WHERE (((pro_empdet.idfun) = 6)) " & nSQLId & nSQLTmp _
                    + vbCr + " ORDER BY [pla_empleados].[apepat] & ' ' & [pla_empleados].[apemat] & ' ' & [pla_empleados].[nom];"
                nTitulo = "Buscando Personal"
            
            Case 5 ' mostrar para ingreso de datos
                pHabilitarBotonEditor True
                Exit Sub
            
            Case Else
                Exit Sub
        End Select
    Else
        If Col <> 3 Then Exit Sub
        ReDim xCampos(4, 3) As String
        xCampos(0, 0) = "Descripción":  xCampos(0, 1) = "nombres": xCampos(0, 2) = "4500":  xCampos(0, 3) = "C"
        xCampos(1, 0) = "Peso":         xCampos(1, 1) = "peso":    xCampos(1, 2) = "800":   xCampos(1, 3) = "N"
        xCampos(2, 0) = "Abrev":        xCampos(2, 1) = "abrev":   xCampos(2, 2) = "700":   xCampos(2, 3) = "C"
        xCampos(3, 0) = "Id":           xCampos(3, 1) = "id":      xCampos(3, 2) = "500":   xCampos(3, 3) = "N"

        nTitulo = "Buscando Contenedor"
        nSQL = "SELECT pro_pesotara.id, pro_pesotara.descripcion AS nombres, pro_pesotara.peso, pro_pesotara.abrev " _
            + vbCr + " FROM mae_unidades INNER JOIN pro_pesotara ON mae_unidades.id = pro_pesotara.idundori "
        ' ojo idunddes =2::kilos
    End If
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, "nombres", "nombres", Principio, ""

    If xRs.State = 0 Then GoTo SALIR
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo SALIR
    
    Agregando = True
    If Index = 0 Then
        ' eliminar el registro que esta reemplazando
        If RstGrDetalle.RecordCount <> 0 Then
            RstGrDetalle.MoveFirst
            Do While Not RstGrDetalle.EOF
                If NulosN(RstGrDetalle("idemp")) = NulosN(fg(0).TextMatrix(fg(0).Row, 4)) Then
                    RstGrDetalle.Delete
                    Exit Do
                End If
                RstGrDetalle.MoveNext
            Loop
        End If
        ' agregando los datos al rst temporal
        'CARGAR_RST_TMP RstGrDetalle, xRs, , , True
        If xRs.State = 1 Then
            RstGrDetalle.AddNew
            RstGrDetalle("codigo") = xRs("codigo")
            RstGrDetalle("idemp") = xRs("idemp")
            RstGrDetalle("idgrupo") = xRs("idgrupo")
            RstGrDetalle("nombres") = xRs("nombres")
            RstGrDetalle("cant") = xRs("cant")
            RstGrDetalle("cantbrut") = xRs("cantbrut")
            RstGrDetalle("activo") = xRs("activo")
        End If
        
        If chkAuto.Value = 1 Then
            pRecalcular
        Else
            pCargarDatos
        End If
    Else
        ' aplicando filtro
        RstGrDetalleTara.Filter = "codigo = " & mCodigoPrincipal & " and tipo=1 and idemp=" & NulosN(fg(0).TextMatrix(fg(0).Row, 4))
        If RstGrDetalleTara.RecordCount <> 0 Then RstGrDetalleTara.MoveFirst
        RstGrDetalleTara.Find "item = " & NulosN(fg(1).TextMatrix(Row, 7))
        If RstGrDetalleTara.EOF = False And RstGrDetalleTara.BOF = False Then
            RstGrDetalleTara("abrev") = NulosC(xRs("abrev"))
            RstGrDetalleTara("idpeso") = NulosN(xRs("id"))
            RstGrDetalleTara("pesouni") = NulosN(xRs("peso"))
        End If
        ' actualizar el grid
        fg(1).TextMatrix(Row, 3) = NulosC(xRs("abrev"))
        fg(1).TextMatrix(Row, 6) = NulosN(xRs("id"))     ' idpeso
        fg(1).TextMatrix(Row, 4) = NulosN(xRs("peso"))
        Agregando = False
        fg_CellChanged 1, Row, 1
        Agregando = True
    End If
    
    Agregando = False
    Set xRs = Nothing
    Exit Sub

SALIR:
    Set xRs = Nothing
    Agregando = False
End Sub

Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If Col = 2 Then
        If validar_letras(KeyAscii) = False Then
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub Fg_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo error
    If Index = 0 Then
        If KeyCode = 114 Or KeyCode = vbKeyInsert Then cmd_Click 0 'F3 = Agregar Item
        If KeyCode = 115 Or KeyCode = vbKeyDelete Then cmd_Click 1         'F4 = Eliminar Item
    Else
        If KeyCode = 114 Or KeyCode = vbKeyInsert Then CmdEditor_Click 0  'F3 = Agregar Item
        If KeyCode = 115 Or KeyCode = vbKeyDelete Then CmdEditor_Click 1         'F4 = Eliminar Item
    End If
    'Exit Sub
    
    '************************************************************************************************
    If Index = 1 Then Exit Sub
    If QueHace = 3 Then Exit Sub
    If fg(0).Col <> 1 Then Exit Sub
    
    If KeyCode <> vbKeySpace And KeyCode <> 13 Then Exit Sub
    
    ' ir al primer registro
    If RstGrDetalle.RecordCount = 0 Then Exit Sub
    RstGrDetalle.MoveFirst
    ' buscar el personal para actualizar si esta activado o no
    RstGrDetalle.Find "idemp=" & NulosN(fg(0).TextMatrix(fg(0).Row, 4))
    If RstGrDetalle.EOF = False And RstGrDetalle.BOF = False Then
        RstGrDetalle("activo") = fg(0).TextMatrix(fg(0).Row, 1)
        ' refrescar los calculos
        If chkAuto.Value = 1 Then
            pRecalcular
        Else
            fg(0).TextMatrix(fg(0).Row, 3) = 0
            fg(0).TextMatrix(fg(0).Row, 5) = 0
            RstGrDetalle("cant") = 0
            RstGrDetalle("cantbrut") = 0
            ' eliminando el registro de taras
            RstGrDetalleTara.Filter = "codigo = " & mCodigoPrincipal & " and tipo=1 and idemp=" & NulosN(fg(0).TextMatrix(fg(0).Row, 4))
            If RstGrDetalleTara.RecordCount <> 0 Then RstGrDetalleTara.MoveFirst
            Do While Not RstGrDetalleTara.EOF
                RstGrDetalleTara.Delete
                RstGrDetalleTara.MoveNext
            Loop
        End If
    End If
    '************************************************************************************************
error:
    SHOW_ERROR Me.Name, "Fg_KeyUp (" & Index & ")"
End Sub

Private Sub Fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace = 3 Then Exit Sub
    If Index = 0 And Button = 2 Then PopupMenu Menu1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift <> 0 Then Exit Sub
    If KeyCode = vbKeyEscape Then
        If FraEditor.Visible = True Then CmdEditor_Click 2
    ElseIf KeyCode = vbKeyF6 Then
        If BuscarFrm("FrmControlTarea1", True, False) = True Then
            If FrmControlTarea1.FraEditor.Visible = False Then
                FrmControlTarea1.Fg1.SetFocus
                Exit Sub
            Else
                If FrmControlTarea1.fg(1).Rows > 1 Then
                    FrmControlTarea1.fg(1).Row = 1
                    FrmControlTarea1.fg(1).Col = 1
                    FrmControlTarea1.fg(1).SetFocus
                Else
                    FrmControlTarea1.CmdEditor(0).SetFocus
                End If
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    iniciarCampos
End Sub

'*****************************************************************************************************
'* Nombre           : pRegistroAdd
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* Paranetros       : NOMBRE    |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Index     |  Integer    |  ESPECIFICA EL INDICE DEL CONTROL Fg1
'* Devuelve         :
'*****************************************************************************************************
Private Sub pRegistroAdd(Index As Integer)
    ' 0 agregar personal; 1 agregar peso-taras
    Dim mCol%
    Dim fInsertar As Boolean
    Agregando = True
    If Index = 0 Then
        If fg(Index).Rows > fg(Index).FixedRows Then
            If NulosN(fg(Index).TextMatrix(fg(Index).Rows - 1, 4)) = 0 Then   ' idEmpleado
                MsgBox "Seleccione el Personal", vbExclamation, xTitulo
            Else
                fInsertar = True
            End If
        Else
            fInsertar = True
        End If
        mCol = 2
    Else
        fInsertar = True
        mCol = 1
    End If
    
    If fInsertar = True Then fg(Index).AddItem ""
    
    fg(Index).Row = fg(Index).Rows - 1
    fg(Index).Col = mCol
    
    ' cargar el buscador por defecto
    If fInsertar = True Then
        If Index = 0 Then
            Fg_CellButtonClick 0, fg(0).Rows - 1, 2
        Else
            ' agregando el registro
            mRowAdd = mRowAdd + 1
            RstGrDetalleTara.AddNew
            RstGrDetalleTara("codigo") = mCodigoPrincipal
            RstGrDetalleTara("tipo") = 1
            RstGrDetalleTara("idemp") = fg(0).TextMatrix(fg(0).Row, 4)
            RstGrDetalleTara("item") = mRowAdd
            fg(1).TextMatrix(fg(1).Rows - 1, 7) = mRowAdd
            
            ' colocar el ultimo peso-tara seleccionado
            If NulosN(RstGrDetalleTara("idpeso")) = 0 And fg(1).Rows > 2 Then
                RstGrDetalleTara("idpeso") = NulosN(fg(Index).TextMatrix(fg(Index).Rows - 2, 6))
                RstGrDetalleTara("pesouni") = NulosN(fg(Index).TextMatrix(fg(Index).Rows - 2, 4))
                RstGrDetalleTara("abrev") = NulosC(fg(Index).TextMatrix(fg(Index).Rows - 2, 3))
                RstGrDetalleTara("cantidad") = NulosN(fg(Index).TextMatrix(fg(Index).Rows - 2, 2))
            Else
                RstGrDetalleTara("idpeso") = NulosN(FrmControlTarea1.lbl_cod(2).Caption)
                RstGrDetalleTara("pesouni") = NulosN(FrmControlTarea1.lblPesoTara(0).Caption)
                RstGrDetalleTara("abrev") = NulosC(FrmControlTarea1.lblPesoTara(1).Caption)
                RstGrDetalleTara("cantidad") = 1
            End If
            
            fg(Index).TextMatrix(fg(Index).Row, 2) = NulosN(RstGrDetalleTara("cantidad"))
            fg(Index).TextMatrix(fg(Index).Row, 3) = NulosC(RstGrDetalleTara("abrev"))
            fg(Index).TextMatrix(fg(Index).Row, 4) = NulosN(RstGrDetalleTara("pesouni"))
            fg(Index).TextMatrix(fg(Index).Row, 6) = NulosN(RstGrDetalleTara("idpeso"))
        End If
    End If
    fg(Index).SetFocus
    Agregando = False
End Sub

'*****************************************************************************************************
'* Nombre           : pRegistroDel
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* Paranetros       : NOMBRE    |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Index     |  Integer    |  ESPECIFICA EL INDICE DEL CONTROL Fg1
'* Devuelve         :
'*****************************************************************************************************
Private Sub pRegistroDel(Index As Integer)
    If fg(Index).Row < 1 Then
        MsgBox "Elija el registro correcto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(Index).SetFocus
        Exit Sub
    End If
    
    If fg(Index).Rows = 1 Then
        MsgBox "No hay Registro para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        fg(Index).SetFocus
        Exit Sub
    End If
    If MsgBox("Seguro desea eliminar el registro", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
    
    If Index = 0 Then
        ' eliminar el registro que esta eliminando
        If RstGrDetalle.RecordCount <> 0 Then RstGrDetalle.MoveFirst
        Do While Not RstGrDetalle.EOF
            If RstGrDetalle.RecordCount = 0 Then Exit Do
            If NulosN(RstGrDetalle("idemp")) = NulosN(fg(0).TextMatrix(fg(0).Row, 4)) Then
                RstGrDetalle.Delete
                Exit Do
            End If
            RstGrDetalle.MoveNext
        Loop
        
        ' eliminando el registro de taras
        RstGrDetalleTara.Filter = "codigo = " & mCodigoPrincipal & " and tipo=1 and idemp=" & NulosN(fg(0).TextMatrix(fg(0).Row, 4))
        If RstGrDetalleTara.RecordCount <> 0 Then RstGrDetalleTara.MoveFirst
        Do While Not RstGrDetalleTara.EOF
            RstGrDetalleTara.Delete
            RstGrDetalleTara.MoveNext
        Loop
        
        ' recalcular de nuevo
        If chkAuto.Value = 1 Then pRecalcular
    Else
        '------------------------------------------------------------------
        '--aplicando filtro
        RstGrDetalleTara.Filter = "codigo = " & mCodigoPrincipal & " and tipo=1 and idemp=" & NulosN(fg(0).TextMatrix(fg(0).Row, 4))
        If RstGrDetalleTara.RecordCount <> 0 Then RstGrDetalleTara.MoveFirst
        RstGrDetalleTara.Find "item = " & NulosN(fg(1).TextMatrix(fg(1).Row, 7))
        If RstGrDetalleTara.EOF = False And RstGrDetalleTara.BOF = False Then
            RstGrDetalleTara.Delete
        End If
        '------------------------------------------------------------------
    End If
    
    fg(Index).RemoveItem fg(Index).Row
    
    If Index = 1 Then
        lblTotal(1).Caption = Format(GRID_SUMAR_COL(fg(1), 1), FORMAT_MONTO)
        lblTotal(3).Caption = Format(GRID_SUMAR_COL(fg(1), 5), FORMAT_MONTO)
    End If
    
    If fg(Index).Rows > 1 Then
        fg(Index).Row = fg(Index).Rows - 1
        fg(Index).Col = 1
        fg(Index).SetFocus
    Else
        If Index = 0 Then
            Cmd(0).SetFocus
        Else
            CmdEditor(0).SetFocus
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarDatos
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA DATOS DEL RECORDSET RstGrDetalle EN EL CONTROL Fg(0)
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarDatos()
    fg(0).Rows = 1
    Agregando = True
    With RstGrDetalle
        If .State = 0 Then Exit Sub
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            fg(0).Rows = fg(0).Rows + 1
            fg(0).TextMatrix(fg(0).Rows - 1, 1) = NulosN(.Fields("activo"))
            fg(0).TextMatrix(fg(0).Rows - 1, 2) = NulosC(.Fields("nombres"))
            fg(0).TextMatrix(fg(0).Rows - 1, 3) = NulosN(.Fields("cantbrut"))
            fg(0).TextMatrix(fg(0).Rows - 1, 5) = NulosN(.Fields("cant"))
            fg(0).TextMatrix(fg(0).Rows - 1, 4) = NulosN(.Fields("idemp"))
            fg(0).TextMatrix(fg(0).Rows - 1, 9) = Format(NulosC(.Fields("horini")), FORMAT_HORA_SIN_SEGUNDO)
            fg(0).TextMatrix(fg(0).Rows - 1, 10) = Format(NulosC(.Fields("horfin")), FORMAT_HORA_SIN_SEGUNDO)
            .MoveNext
        Loop
    End With
    
    ' aplicando el orden a la lista de datos
    GRID_ORDENAR fg(0), 1, 2
    Agregando = False
End Sub

'*****************************************************************************************************
'* Nombre           : pRecalcular
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : Redistribuir la Cantidad en funcion a la cantidad de personal, Cantidades
'*                    distribuidos al personal
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pRecalcular()
    Dim nTotalPersonal As Long
    Dim sTotalxPersonal As Double
    
    If sCantidad <> 0 Then
        If RstGrDetalle.RecordCount > 0 Then
            RstGrDetalle.MoveFirst
            Do While Not RstGrDetalle.EOF
                If NulosN(RstGrDetalle("activo")) = -1 Then nTotalPersonal = nTotalPersonal + 1
                RstGrDetalle.MoveNext
            Loop
        End If
        
        ' obtener la cantidad que le corresponde por cada integrante
        If nTotalPersonal > 0 Then
            sTotalxPersonal = sCantidad / nTotalPersonal
        Else
            sTotalxPersonal = 0
        End If
    End If
    
    ' actualizar las cantidades a los integrantes
    If RstGrDetalle.RecordCount > 0 Then
        RstGrDetalle.MoveFirst
        Do While Not RstGrDetalle.EOF
            If NulosN(RstGrDetalle("activo")) <> 0 Then
                RstGrDetalle("cant") = sTotalxPersonal
            Else
                RstGrDetalle("cant") = 0
                RstGrDetalle("horini") = Null
                RstGrDetalle("horfin") = Null
            End If
            RstGrDetalle.MoveNext
        Loop
    End If
    
    pCargarDatos
    fCalculoAuto = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    fDesactivarAuto = False
    Set RstGrDetalle = Nothing
    Set RstGrDetalleTara = Nothing
End Sub

Private Sub Menu1_1_Click()
    cmd_Click 0
End Sub

Private Sub Menu1_3_Click()
    cmd_Click 2
End Sub

'*****************************************************************************************************
'* Nombre           : pHabilitarBotonEditor
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : Registrar las cantidades del personal Ej. peso bruto y cantidades Mostrar/Ocultar
'*                    las opciones del detalle personal
'* Paranetros       : NOMBRE    |  TIPO     |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    band      |  Boolean  |  band= puede ser true o false
'* Devuelve         :
'*****************************************************************************************************
Private Sub pHabilitarBotonEditor(band As Boolean)
    ' true muestra el ingreso de datos
    If band = True Then
        fg(0).Enabled = False
        FraEditor.Top = 30
        FraEditor.Left = 30
        LblTituloFrame.Caption = "Registros: " & fg(0).TextMatrix(fg(0).Row, 2)
    Else
        fg(0).Enabled = True
    End If
    FraEditor.Visible = band
    habilitar Cmd, Not band
    
    ' si es true cargar los datos
    Agregando = True
    If band = True Then
        fg(1).Rows = 1
        With RstGrDetalleTara
            '--filtrar los registros solo del personal seleccionado
            .Filter = "codigo = " & mCodigoPrincipal & " and idemp=" & NulosN(fg(0).TextMatrix(fg(0).Row, 4))
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                fg(1).Rows = fg(1).Rows + 1
                fg(1).TextMatrix(fg(1).Rows - 1, 1) = NulosN(.Fields("pesobrut"))
                fg(1).TextMatrix(fg(1).Rows - 1, 2) = NulosN(.Fields("cantidad"))
                fg(1).TextMatrix(fg(1).Rows - 1, 3) = NulosC(.Fields("abrev"))
                fg(1).TextMatrix(fg(1).Rows - 1, 4) = NulosN(.Fields("pesouni"))
                fg(1).TextMatrix(fg(1).Rows - 1, 5) = NulosN(.Fields("pesonet"))
                fg(1).TextMatrix(fg(1).Rows - 1, 6) = NulosN(.Fields("idpeso"))
                fg(1).TextMatrix(fg(1).Rows - 1, 7) = NulosN(.Fields("item")) '--identificador de fila
                
                .MoveNext
            Loop
        End With
        
        If fg(1).Rows > 1 Then
            fg(1).Row = fg(1).Rows - 1
            fg(1).Col = 1
            fg(1).SetFocus
        Else
            CmdEditor(0).SetFocus
        End If
        
        lblTotal(1).Caption = Format(GRID_SUMAR_COL(fg(1), 1), FORMAT_MONTO)
        lblTotal(3).Caption = Format(GRID_SUMAR_COL(fg(1), 5), FORMAT_MONTO)
    Else
        If NulosN(lblTotal(3).Caption) <> 0 Then
            ' acumular las cantidades
            fg(0).TextMatrix(fg(0).Row, 3) = GRID_SUMAR_COL(fg(1), 1)
            fg(0).TextMatrix(fg(0).Row, 5) = GRID_SUMAR_COL(fg(1), 5)
            ' actualizar el rsttemporal
            ' aplicando filtro
            RstGrDetalle.Filter = "codigo = " & mCodigoPrincipal
            If RstGrDetalle.RecordCount <> 0 Then RstGrDetalle.MoveFirst
            RstGrDetalle.Find "idemp= " & NulosN(fg(0).TextMatrix(fg(0).Row, 4))
            If RstGrDetalle.EOF = False And RstGrDetalle.BOF = False Then
                RstGrDetalle("cantbrut") = NulosN(fg(0).TextMatrix(fg(0).Row, 3))
                RstGrDetalle("cant") = NulosN(fg(0).TextMatrix(fg(0).Row, 5))
            End If
            
            fDesactivarAuto = True
            FrmControlTarea1.Fg1.TextMatrix(FrmControlTarea1.Fg1.Row, 8) = GRID_SUMAR_COL(fg(0), 5)
            fDesactivarAuto = False
        End If
        fg(0).Row = fg(0).Row
        fg(0).Col = 5
        fg(0).SetFocus
    End If
    Agregando = False
End Sub

Private Sub CmdEditor_Click(Index As Integer)
    Select Case Index
        Case 0 ' agregar
            pRegistroAdd 1
        Case 1 ' eliminar
            pRegistroDel 1
        Case 2 ' cancelar
            pHabilitarBotonEditor False
    End Select
End Sub

Private Sub Option1_Click()
    If QueHace = 3 Then Exit Sub
    seleccionartodos True
End Sub

'Private Sub Option1_Validate(Cancel As Boolean)
'    If Option1.Value = True Then
'        seleccionartodos True
'    Else
'        seleccionartodos False
'    End If
'End Sub

Private Sub Option2_Click()
    If QueHace = 3 Then Exit Sub
    seleccionartodos False
End Sub

Private Sub pic_Click()
    CmdEditor_Click 2
End Sub


