VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmGeneraHoja 
   Caption         =   "Produccion - Hoja de Ruta"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4200
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option2 
      Caption         =   "Producto"
      Height          =   210
      Left            =   1665
      TabIndex        =   16
      Top             =   180
      Width           =   1320
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Materia Prima"
      Height          =   210
      Left            =   75
      TabIndex        =   15
      Top             =   180
      Width           =   1320
   End
   Begin VB.TextBox TxtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4785
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "TxtTotal"
      Top             =   2895
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   45
      TabIndex        =   8
      Top             =   3375
      Width           =   6810
      Begin VB.CommandButton Command2 
         Caption         =   "&Hoja de Tareas"
         Height          =   390
         Left            =   2745
         TabIndex        =   17
         Top             =   255
         Width           =   1200
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   3975
         TabIndex        =   10
         Top             =   255
         Width           =   1200
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Imprimir Hoja Ruta"
         Height          =   390
         Left            =   1515
         TabIndex        =   9
         Top             =   255
         Width           =   1200
      End
   End
   Begin VB.CommandButton CmdBusMatPri 
      Height          =   240
      Left            =   6630
      Picture         =   "FrmGeneraHoja.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   555
      Width           =   225
   End
   Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmi 
      Height          =   300
      Left            =   1290
      TabIndex        =   2
      Top             =   2970
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.TextBox TxtCantidad 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1290
      TabIndex        =   1
      Text            =   "TxtCantidad"
      Top             =   855
      Width           =   1200
   End
   Begin VB.TextBox TxtProducto 
      Height          =   300
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "TxtProducto"
      Top             =   525
      Width           =   5595
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg2 
      Height          =   1470
      Left            =   45
      TabIndex        =   12
      Top             =   1425
      Width           =   6825
      _cx             =   12039
      _cy             =   2593
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
      BackColorSel    =   64
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
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmGeneraHoja.frx":0132
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Productos"
      Height          =   195
      Left            =   60
      TabIndex        =   14
      Top             =   1185
      Width           =   720
   End
   Begin VB.Label Label11 
      Caption         =   "Total ==>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3810
      TabIndex        =   13
      Top             =   2940
      Width           =   825
   End
   Begin VB.Label LblIdProducto 
      Caption         =   "LblIdProducto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   5580
      TabIndex        =   7
      Top             =   885
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fch. de Emision"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   3015
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   900
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Materia Prima"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   570
      Width           =   960
   End
End
Attribute VB_Name = "FrmGeneraHoja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean

Private Sub CmdAceptar_Click()
    If Option1.Value = True Then
        If NulosC(TxtProducto.Text) = "" Then
            MsgBox "No ha especificado el producto que se va a procesar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtProducto.SetFocus
            Exit Sub
        End If
        
        If NulosN(TxtCantidad.Text) = 0 Then
            MsgBox "No ha especificado la cantidad de producto a procesar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtCantidad.SetFocus
            Exit Sub
        End If
        
        If NulosN(TxtTotal.Text) <> NulosN(TxtCantidad.Text) Then
            MsgBox "El importe a procesar en productos no coincide con la catidad de materia prima ingresada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
            TxtTotal.SetFocus
        End If
    End If
    
    Dim objExcel As Excel.Application
    Dim xLibro As Excel.Workbook
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    Dim xPor As Double
    
    Set objExcel = New Excel.Application
    Set xLibro = objExcel.Workbooks.Open(Trim(App.Path) & "\hojaruta.xls")
    
    With objExcel.ActiveSheet
        .Cells(3, 4).Value = Fg2.TextMatrix(Fg2.Row, 1)
        .Cells(4, 4).Value = Format(Fg2.TextMatrix(Fg2.Row, 2), "0.00")
        .Cells(4, 10).Value = TxtFchEmi.valor
        
        RST_Busq Rst, "SELECT pro_receta.iditem, pro_tareas.descripcion, pro_recetatar.numper, pro_recetatar.factor, pro_recetatar.idtar, pro_receta.id, pro_recetatar.aplpor" _
            & " FROM (pro_receta LEFT JOIN pro_recetatar ON pro_receta.id = pro_recetatar.idrec) LEFT JOIN pro_tareas ON pro_recetatar.idtar = pro_tareas.id " _
            & " Where (((pro_receta.iditem) = " & NulosN(Fg2.TextMatrix(Fg2.Row, 4)) & ") And ((pro_recetatar.numper) <> 0) And ((pro_recetatar.factor) <> 0)) " _
            & " ORDER BY pro_recetatar.orden", xCon

        'Rst.Filter = "idtar = 5"
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            Dim xFil As Integer
            Dim xTiempo As Double
            Dim xHorEst As String
            xFil = 8
            For A = 1 To Rst.RecordCount
                .Cells(xFil, 3).Value = Rst("descripcion")
                .Cells(xFil, 7).Value = Rst("numper")
                
                If Rst("aplpor") <> 0 Then
                    xPor = 0
                    xPor = (Rst("aplpor") / 100)
                    xPor = NulosN(Fg2.TextMatrix(Fg2.Row, 2)) * xPor
                    ' SI ES MATERIA PRIMA CALCULAMOS EL TOTAL A PRODUCIR DEL PRODUCTO EN FUNCION AL RENDIMIENTO
                    xTiempo = (NulosN(Rst("factor")) * xPor)
                    xTiempo = xTiempo / Rst("numper")
                Else
                    xTiempo = (NulosN(Rst("factor")) * NulosN(Fg2.TextMatrix(Fg2.Row, 2)))
                    xTiempo = xTiempo / Rst("numper")
                End If
                xHorEst = Format(Int(xTiempo), "00")
                xHorEst = xHorEst & ":" & Format(((xTiempo * 60) Mod 60), "00")
                
                '.Cells(xFil, 8).Value = xHorEst
                
                Rst.MoveNext
                
                If Rst.EOF = True Then Exit For
                xFil = xFil + 1
            Next A
        End If
        
        Set Rst = Nothing

        RST_Busq Rst, "SELECT pro_receta.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro " _
            & " FROM ((pro_receta LEFT JOIN pro_recetains ON pro_receta.id = pro_recetains.idrec) LEFT JOIN alm_inventario ON pro_recetains.iditem = alm_inventario.id) " _
            & " LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id Where (((pro_receta.iditem) = " & NulosN(Fg2.TextMatrix(Fg2.Row, 4)) & ") And ((pro_receta.prirec) = 1))" _
            & " ORDER BY alm_inventario.descripcion", xCon

        
        
        'SELECT pro_receta.iditem, alm_inventario.descripcion, mae_unidades.abrev, pro_recetains.canpro" _
            & " FROM ((pro_receta LEFT JOIN pro_recetains ON pro_receta.id = pro_recetains.idrec) LEFT JOIN alm_inventario " _
            & " ON pro_recetains.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON pro_recetains.idunimed = mae_unidades.id " _
            & " Where (((pro_receta.iditem) = " & NulosN(Fg2.TextMatrix(Fg2.Row, 4)) & ")) ORDER BY alm_inventario.descripcion", xCon

        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst

            xFil = 23
            For A = 1 To Rst.RecordCount
                .Cells(xFil, 3).Value = Rst("descripcion")
                .Cells(xFil, 7).Value = Rst("abrev")
                
                If Option1.Value = True Then
                    xPor = 0
                    xPor = (NulosN(Fg2.TextMatrix(Fg2.Row, 5)) / 100)
                    xPor = NulosN(Fg2.TextMatrix(Fg2.Row, 2)) * xPor
                    ' SI ES MATERIA PRIMA CALCULAMOS EL TOTAL A PRODUCIR DEL PRODUCTO EN FUNCION AL RENDIMIENTO
                    .Cells(xFil, 8).Value = Format(Rst("canpro") * xPor, "0.000000")
                Else
                    .Cells(xFil, 8).Value = Format(Rst("canpro") * Fg2.TextMatrix(Fg2.Row, 2), "0.000000")
                End If

                Rst.MoveNext
                If Rst.EOF = True Then Exit For
                xFil = xFil + 1
            Next A
        End If
    End With
    
    objExcel.Visible = True
    Set objExcel = Nothing
    Exit Sub
End Sub

Private Sub CmdBusMatPri_Click()
    'If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Uni. Med.":     xCampos(2, 1) = "abrev":          xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
    
    
    xform.SQLCad = "SELECT alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.activo,  alm_inventario.id" _
        & " FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed " _
        & " Where (((alm_inventario.tippro) = 1) And ((alm_inventario.activo) = -1)) " _
        & " ORDER BY alm_inventario.descripcion"
    xform.titulo = "Buscando Productos"
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtProducto.Text = xRs("descripcion")
            LblIdProducto.Caption = xRs("id")
            
            Dim Rst As New ADODB.Recordset
            Dim A As Integer
            RST_Busq Rst, "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, pro_redimiento.rend " _
                & " FROM pro_redimiento LEFT JOIN alm_inventario ON pro_redimiento.idpro = alm_inventario.id " _
                & " WHERE (((pro_redimiento.iditem)=" & xRs("id") & "))", xCon
            
            If Rst.RecordCount <> 0 Then
                Fg2.Rows = 1
                
                Rst.MoveFirst
                For A = 1 To Rst.RecordCount
                    Fg2.Rows = Fg2.Rows + 1
                    
                    Fg2.TextMatrix(A, 1) = Rst("descripcion")
                    Fg2.TextMatrix(A, 2) = ""
                    Fg2.TextMatrix(A, 4) = Rst("id")
                    Fg2.TextMatrix(A, 5) = Rst("rend")
                    Rst.MoveNext
                    If Rst.EOF = True Then Exit For
                Next A
                
                If Rst.RecordCount = 1 Then
                    Fg2.Editable = flexEDNone
                    Fg2.TextMatrix(1, 2) = TxtCantidad.Text
                    Fg2.TextMatrix(1, 3) = 1
                Else
                    Fg2.Editable = flexEDKbdMouse
                End If
            End If
            TxtCantidad.SetFocus
            
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    If Fg2.Rows = 1 Then
        MsgBox "No se ha especificado un producto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Dim objExcel As Excel.Application
    Dim xLibro As Excel.Workbook
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    
    Set objExcel = New Excel.Application
    Set xLibro = objExcel.Workbooks.Open(Trim(App.Path) & "\hojaTareas.xls")

    With objExcel.ActiveSheet
        .Cells(1, 4).Value = Fg2.TextMatrix(Fg2.Row, 1)
        .Cells(1, 9).Value = Format(Fg2.TextMatrix(Fg2.Row, 2), "0.00")
        .Cells(2, 4).Value = TxtFchEmi.valor
    End With
    
    objExcel.Visible = True
    Set objExcel = Nothing
    Exit Sub

End Sub

Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Uni. Med.":     xCampos(2, 1) = "abrev":          xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
    
    
    xform.SQLCad = "SELECT alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.activo,  alm_inventario.id" _
        & " FROM mae_unidades RIGHT JOIN alm_inventario ON mae_unidades.id = alm_inventario.idunimed " _
        & " Where (((alm_inventario.tippro) = 3) And ((alm_inventario.activo) = -1)) " _
        & " ORDER BY alm_inventario.descripcion"
    
    xform.titulo = "Buscando Productos"
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            Fg2.TextMatrix(Fg2.Row, 1) = xRs("descripcion")
            Fg2.TextMatrix(Fg2.Row, 4) = xRs("id")
            
'            TxtProducto.Text = xRs("descripcion")
'            LblIdProducto.Caption = xRs("id")
'
'            Dim Rst As New ADODB.Recordset
'            Dim A As Integer
'            RST_Busq Rst, "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.descripcion, pro_redimiento.rend " _
'                & " FROM pro_redimiento LEFT JOIN alm_inventario ON pro_redimiento.idpro = alm_inventario.id " _
'                & " WHERE (((pro_redimiento.iditem)=" & xRs("id") & "))", xCon
'
'            If Rst.RecordCount <> 0 Then
'                Fg2.Rows = 1
'
'                Rst.MoveFirst
'                For A = 1 To Rst.RecordCount
'                    Fg2.Rows = Fg2.Rows + 1
'
'                    Fg2.TextMatrix(A, 1) = Rst("descripcion")
'                    Fg2.TextMatrix(A, 2) = ""
'                    Fg2.TextMatrix(A, 4) = Rst("id")
'                    Fg2.TextMatrix(A, 5) = Rst("rend")
'                    Rst.MoveNext
'                    If Rst.EOF = True Then Exit For
'                Next A
'
'                If Rst.RecordCount = 1 Then
'                    Fg2.Editable = flexEDNone
'                    Fg2.TextMatrix(1, 2) = TxtCantidad.Text
'                    Fg2.TextMatrix(1, 3) = 1
'                Else
'                    Fg2.Editable = flexEDKbdMouse
'                End If
'            End If
'            TxtCantidad.SetFocus
            
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
    
End Sub

Private Sub Fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Fg2.Col = 2 Then
        Fg2.TextMatrix(Row, Col) = Format(Fg2.TextMatrix(Row, Col), "0.00")
        
        TxtTotal.Text = Format(GRID_SUMAR_COL(Fg2, 2, 1, Fg2.Rows - 1), "0.00")
        If NulosN(Fg2.TextMatrix(Row, Col)) <> 0 Then
            Fg2.TextMatrix(Row, 3) = 1
        Else
            Fg2.TextMatrix(Row, 3) = 0
        End If
    End If
    If Fg2.Col = 3 Then
        If NulosN(Fg2.TextMatrix(Row, Col)) = 0 Then
            Fg2.TextMatrix(Row, 2) = ""
        End If
        TxtTotal.Text = Format(GRID_SUMAR_COL(Fg2, 2, 1, Fg2.Rows - 1), "0.00")
    End If
End Sub

Private Sub Fg2_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    'If QueHace = 3 Then KeyAscii = 0
    
    If KeyAscii = 13 Then Exit Sub
    ' validar los caracteres que se ingresan
    Select Case Col
        Case 1, 3
            KeyAscii = 0
            
        Case 2
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End Select
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        Blanquea
        TxtProducto.SetFocus
        TxtFchEmi.valor = Date
    End If
End Sub

Sub Blanquea()
    TxtProducto.Text = ""
    TxtCantidad.Text = ""
    TxtFchEmi.valor = ""
    LblIdProducto.Caption = ""
    TxtTotal.Text = ""
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    Option1.Value = True
    Option1_Click
    Fg2.ColWidth(4) = 0
    Fg2.ColWidth(5) = 0
    Fg2.Rows = 1
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
        TxtProducto.Enabled = True
        TxtCantidad.Enabled = True
        CmdBusMatPri.Enabled = True
        
        Fg2.ColComboList(1) = ""
        Fg2.Editable = flexEDNone
        TxtProducto.BackColor = &H80000005
        TxtCantidad.BackColor = &H80000005
    End If
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then
        TxtProducto.Text = ""
        TxtCantidad.Text = ""
        LblIdProducto.Caption = ""
        
        Fg2.Rows = 1
        TxtTotal.Text = ""
        
        Fg2.ColComboList(1) = "|..."
        Fg2.Editable = flexEDKbdMouse
        
        Fg2.Rows = Fg2.Rows + 1
        TxtProducto.Enabled = False
        TxtCantidad.Enabled = False
        CmdBusMatPri.Enabled = False
        TxtProducto.BackColor = &H8000000F
        TxtCantidad.BackColor = &H8000000F
    End If
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub TxtCantidad_Validate(Cancel As Boolean)
    If NulosN(TxtCantidad.Text) = 0 Then
        TxtCantidad.Text = "0.00"
        Exit Sub
    Else
        TxtCantidad.Text = Format(TxtCantidad.Text, "0.00")
        If Fg2.Rows = 2 Then
            Fg2.TextMatrix(1, 2) = TxtCantidad.Text
            TxtTotal.Text = Format(GRID_SUMAR_COL(Fg2, 2, 1, Fg2.Rows - 1), "0.00")
        End If
    End If
End Sub

Private Sub TxtProducto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtProducto_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMatPri_Click
    End If
End Sub
