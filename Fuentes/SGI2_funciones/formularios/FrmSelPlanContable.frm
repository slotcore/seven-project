VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Begin VB.Form FrmSelPlanContable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Seleccion de Plan de Cuentas"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   4830
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   9570
      _cx             =   16880
      _cy             =   8520
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmSelPlanContable.frx":0000
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
      Height          =   735
      Left            =   30
      TabIndex        =   1
      Top             =   4845
      Width           =   9585
      Begin VB.CommandButton CmdSubNiv 
         Height          =   450
         Left            =   180
         Picture         =   "FrmSelPlanContable.frx":00AD
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   195
         Width           =   1485
      End
      Begin VB.CommandButton CmdBajaNiv 
         Height          =   450
         Left            =   2400
         Picture         =   "FrmSelPlanContable.frx":2B23
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   195
         Width           =   1485
      End
      Begin VB.TextBox TxtNivAct 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "TxtNivAct"
         Top             =   210
         Width           =   645
      End
      Begin VB.CommandButton CmdSalir 
         Height          =   450
         Left            =   8205
         Picture         =   "FrmSelPlanContable.frx":5599
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   1170
      End
      Begin VB.CommandButton CmdAceptar 
         Height          =   450
         Left            =   6990
         Picture         =   "FrmSelPlanContable.frx":73EB
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   1170
      End
   End
End
Attribute VB_Name = "FrmSelPlanContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public RstSele As New ADODB.Recordset
Dim Rst As New ADODB.Recordset
Dim SeEjecuto As Boolean
Dim Nivel As Integer

Dim xLong As Integer
Dim Mostrando As Boolean

Private Sub CmdAceptar_Click()
    PreparaRst
    RstSele.AddNew
    RstSele("id") = Fg1.TextMatrix(Fg1.Row, 4)
    RstSele("cuenta") = Fg1.TextMatrix(Fg1.Row, 1)
    RstSele("descripcion") = Fg1.TextMatrix(Fg1.Row, 2)
    RstSele.Update
    CmdSalir_Click
End Sub

Private Sub CmdBajaNiv_Click()
    If Val(Fg1.TextMatrix(Fg1.Row, 5)) = 0 Then
        MsgBox "La cuenta contable seleccionada no contiene divicionarias", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    MuestraNivel 1
End Sub

Private Sub CmdSalir_Click()
    Set Rst = Nothing
    Unload Me
End Sub

Private Sub CmdSubNiv_Click()
     If Nivel = 1 Then
        MsgBox "Esta en el nivel superior", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
     End If
     MuestraNivel 2
End Sub

Private Sub Fg1_DblClick()
    If Fg1.TextMatrix(Fg1.Row, 5) = 0 Then
        CmdAceptar_Click
        Exit Sub
    End If
    MuestraNivel 1
End Sub

Sub MuestraNivel(Accion As Integer)
    'Accion = 1  Incrementa el nivel
    'Accion = 2  Disminuir el nivel
    
    Dim A As Integer
    Dim xCad As String
    
    If Accion = 1 Then
        Nivel = Nivel + 1
        'NumNivel = NumNivel + 1
        TxtNivAct.Text = Trim(Str(Nivel))
        If Nivel = 2 Then
            xLong = xLong + 2
            RST_Busq Rst, "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id, " _
                & " Len(Trim([con_planctas]![cuenta])) AS numcar, con_planctas.tipo " _
                & " From con_planctas Where (((con_planctas.cuenta) Like '" & Mid(Fg1.TextMatrix(Fg1.Row, 1), 1, 2) & "%') " _
                & " And ((Len(Trim([con_planctas]![cuenta]))) = " & xLong & ")) " _
                & " ORDER BY con_planctas.cuenta", xCon
        End If
        If Nivel >= 3 Then
            xLong = xLong + 3
            RST_Busq Rst, "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id, " _
                & " Len(Trim([con_planctas]![cuenta])) AS numcar, con_planctas.tipo " _
                & " From con_planctas Where (((con_planctas.cuenta) Like '" & Mid(Fg1.TextMatrix(Fg1.Row, 1), 1, xLong - 3) & "%') " _
                & " And ((Len(Trim([con_planctas]![cuenta]))) = " & xLong & ")) " _
                & " ORDER BY con_planctas.cuenta", xCon
        End If
    Else
        Nivel = Nivel - 1
        TxtNivAct.Text = Trim(Str(Nivel))
        If Nivel = 1 Then
            xLong = 2
            RST_Busq Rst, "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id,  con_planctas.tipo, " _
            & " Len(Trim([con_planctas]![cuenta])) AS numcar From con_planctas WHERE (((Len(Trim([con_planctas]![cuenta])))=2)) " _
            & " ORDER BY cuenta", xCon
        End If
        If Nivel = 2 Then
            xLong = xLong - 3
            RST_Busq Rst, "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id, " _
                & " Len(Trim([con_planctas]![cuenta])) AS numcar, con_planctas.tipo " _
                & " From con_planctas Where (((con_planctas.cuenta) Like '" & Mid(Fg1.TextMatrix(Fg1.Row, 1), 1, 2) & "%') " _
                & " And ((Len(Trim([con_planctas]![cuenta]))) = " & 4 & ")) " _
                & " ORDER BY con_planctas.cuenta", xCon
        End If
        If Nivel >= 3 Then
            xLong = xLong - 3
            RST_Busq Rst, "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id, " _
                & " Len(Trim([con_planctas]![cuenta])) AS numcar, con_planctas.tipo " _
                & " From con_planctas Where (((con_planctas.cuenta) Like '" & Mid(Fg1.TextMatrix(Fg1.Row, 1), 1, xLong - 3) & "%') " _
                & " And ((Len(Trim([con_planctas]![cuenta]))) = " & xLong & ")) " _
                & " ORDER BY con_planctas.cuenta", xCon
        End If
    End If

    If Rst.RecordCount = 0 Then
        If Accion = 1 Then Nivel = Nivel - 2
        MsgBox "La cuenta contable seleccionada no contiene divicionarias", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Fg1.Rows = 1
    Mostrando = True
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = Rst("cuenta")
            Fg1.TextMatrix(A, 2) = Rst("descripcion")
            Fg1.TextMatrix(A, 4) = Rst("id")
            Fg1.TextMatrix(A, 5) = Rst("tipo")
            
            If Rst("tipo") = 1 Then
                With Fg1
                    .Select A, 1, A, 3
                    .FillStyle = flexFillRepeat
                    .CellForeColor = &H800000
                    .CellFontBold = True
                End With
            Else
                With Fg1
                    .Select A, 1, A, 3
                    .FillStyle = flexFillRepeat
                    .CellForeColor = &H80000008
                    .CellFontBold = False
                End With
            End If

            
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
    End If
    
    With Fg1
        .Select 1, 1, 1, 3
    End With
    Fg1.SetFocus
    Mostrando = False
End Sub

Private Sub Fg1_EnterCell()
    If Fg1.Rows = 1 Then Exit Sub
    
    If Mostrando = True Then Exit Sub
    
    If Fg1.Col = 3 Then
        If Fg1.TextMatrix(Fg1.Row, 5) = 1 Then
            Fg1.TextMatrix(Fg1.Row, 3) = 0
            Fg1.Editable = flexEDNone
        Else
            Fg1.Editable = flexEDKbdMouse
        End If
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If NulosN(Fg1.TextMatrix(Fg1.Row, 5)) = 0 Then
            CmdAceptar_Click
            Exit Sub
        End If
        
        If KeyAscii = 13 Then MuestraNivel 1
    End If
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 37 Then  'flecha a ala izquierda
        MuestraNivel 2
    End If
    If KeyCode = 39 Then  'flecha a ala izquierda
        MuestraNivel 1
    End If
    
'    If KeyCode = 45 Then  'tecla insert
'        If Fg1.TextMatrix(Fg1.Row, 5) = 1 Then
'            MsgBox "No puede agregar esta cuenta contable, seleccione una de sus divicionarias", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'            Exit Sub
'        End If
'
'        Dim B As Integer
'        If Fg1.TextMatrix(Fg1.Row, 4) = "1" Then Exit Sub
'
''        For B = 1 To Fg2.Rows - 1
''            If Fg1.TextMatrix(Fg1.Row, 1) = Fg2.TextMatrix(B, 1) Then
''                MsgBox "la cuenta contable ya fue agregada a la lista", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
''                Exit Sub
''            End If
''        Next B
'
'
'        'copiamos el plan de cuenta seleccionado
'        RstSele.AddNew
'        RstSele("id") = Fg1.TextMatrix(Fg1.Row, 4)
'        RstSele("cuenta") = Fg1.TextMatrix(Fg1.Row, 1)
'        RstSele("descripcion") = Fg1.TextMatrix(Fg1.Row, 2)
'        RstSele.Update
'    End If
End Sub

Private Sub Fg1_RowColChange()
'    If Mostrando = True Then Exit Sub
'
''    Fg1.BackColorSel = Fg1.BackColor
'    If Fg1.TextMatrix(Fg1.Row, 5) = 1 Then
'        Fg1.ForeColorSel = &H800000 '&HFF&
'    Else
'        Fg1.ForeColorSel = &H80000008  '&HFF&
'    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        Dim A As Integer
        Nivel = 1
        xLong = xLong + 2
        SeEjecuto = True
        Mostrando = True
        TxtNivAct.Text = Trim(Str(Nivel))
        PreparaRst
        
        RST_Busq Rst, "SELECT con_planctas.cuenta, con_planctas.descripcion, con_planctas.id,  con_planctas.tipo, " _
            & " Len(Trim([con_planctas]![cuenta])) AS numcar From con_planctas WHERE (((Len(Trim([con_planctas]![cuenta])))=2)) " _
            & " ORDER BY cuenta", xCon
    
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            For A = 1 To Rst.RecordCount
                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(A, 1) = Rst("cuenta")
                Fg1.TextMatrix(A, 2) = Rst("descripcion")
                Fg1.TextMatrix(A, 4) = Rst("id")
                Fg1.TextMatrix(A, 5) = Rst("tipo")
                
                If Rst("tipo") = 1 Then
                    With Fg1
                        .Select A, 1, A, 3
                        .FillStyle = flexFillRepeat
                        .CellForeColor = &H800000
                        .CellFontBold = True
                    End With
                Else
                    With Fg1
                        .Select A, 1, A, 3
                        .FillStyle = flexFillRepeat
                        .CellForeColor = &H80000008
                        .CellFontBold = False
                    End With
                End If
                
                Rst.MoveNext
                If Rst.EOF = True Then
                    Exit For
                End If
            Next A
            
            With Fg1
                .Select 1, 1, 1, 3
            End With

            Mostrando = False
        End If
    End If
End Sub

Private Sub Form_Load()
    Fg1.Rows = 1
    Fg1.ColWidth(4) = 0
    Fg1.ColWidth(5) = 0
    
    Fg1.ColWidth(3) = 0
    SeEjecuto = False
    xLong = 0
    Fg1.SelectionMode = flexSelectionByRow
End Sub

Sub PreparaRst()
    'Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(3, 3) As String

    xCampos(0, 0) = "id":     xCampos(0, 1) = "N":      xCampos(0, 2) = "3"
    xCampos(1, 0) = "cuenta":       xCampos(1, 1) = "C":      xCampos(1, 2) = "12"
    xCampos(2, 0) = "descripcion":  xCampos(2, 1) = "C":      xCampos(2, 2) = "150"
    
    Set RstSele = xFun.CrearRstTMP(xCampos)
    RstSele.Open
End Sub

