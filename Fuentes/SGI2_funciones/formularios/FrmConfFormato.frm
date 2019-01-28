VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmConfFormato 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seven - Configuración de Libros Contables"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdEjecutar 
      Caption         =   "Establecer por Defecto"
      Height          =   990
      Left            =   9735
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   615
      Width           =   2190
   End
   Begin VB.TextBox TxtLibro 
      Height          =   285
      Left            =   1230
      TabIndex        =   7
      Text            =   "TxtLibro"
      Top             =   60
      Width           =   7290
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   2685
      Left            =   15
      TabIndex        =   0
      Top             =   2145
      Width           =   11895
      _cx             =   20981
      _cy             =   4736
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmConfFormato.frx":0000
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
      Height          =   930
      Left            =   15
      TabIndex        =   1
      Top             =   5235
      Width           =   11925
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Guardar"
         Height          =   720
         Left            =   7035
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   150
         Width           =   810
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "&Guardar"
         Height          =   720
         Left            =   4305
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   150
         Width           =   810
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   720
         Left            =   5130
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   150
         Width           =   810
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   720
         Left            =   5955
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   150
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "7 = Derecha"
         Height          =   195
         Left            =   1440
         TabIndex        =   17
         Top             =   420
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "4 = Izquierda"
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   645
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "1 = Centro"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Alineacion 
         AutoSize        =   -1  'True
         Caption         =   "Alineacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   165
         TabIndex        =   14
         Top             =   165
         Width           =   900
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid fg2 
      Height          =   990
      Left            =   30
      TabIndex        =   4
      Top             =   615
      Width           =   9585
      _cx             =   16907
      _cy             =   1746
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmConfFormato.frx":0189
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
   Begin VSFlex7Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   270
      Left            =   15
      TabIndex        =   19
      Top             =   1905
      Width           =   11895
      _cx             =   20981
      _cy             =   476
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmConfFormato.frx":0216
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   45
      TabIndex        =   11
      Top             =   4950
      Width           =   1020
   End
   Begin VB.Label LblDesc 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblDesc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   300
      Left            =   1230
      TabIndex        =   10
      Top             =   4920
      Width           =   10695
   End
   Begin VB.Label LblIdLibro 
      AutoSize        =   -1  'True
      Caption         =   "LblIdLibro"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   8640
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Libro"
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
      Index           =   2
      Left            =   45
      TabIndex        =   8
      Top             =   90
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Campos del Libro"
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
      Index           =   1
      Left            =   45
      TabIndex        =   6
      Top             =   1680
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Opciones del Libro"
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
      Index           =   0
      Left            =   45
      TabIndex        =   5
      Top             =   405
      Width           =   1605
   End
End
Attribute VB_Name = "FrmConfFormato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean
Dim Tabla1, Tabla2, Tabla3 As String

Private Sub CmdCancel_Click()
    fg2.SelectionMode = flexSelectionByRow
    Fg1.SelectionMode = flexSelectionByRow
    fg2.Enabled = True
    ActivarBotton
    fg2.Editable = flexEDNone
    Fg1.Editable = flexEDNone
End Sub

Private Sub CmdEdit_Click()
    fg2.SelectionMode = flexSelectionFree
    Fg1.SelectionMode = flexSelectionFree
    fg2.Enabled = False
    ActivarBotton
    fg2.Editable = flexEDKbdMouse
    Fg1.Editable = flexEDKbdMouse
End Sub

Sub ActivarBotton()
    CmdEdit.Enabled = Not CmdEdit.Enabled
    CmdGrabar.Enabled = Not CmdGrabar.Enabled
    CmdCancel.Enabled = Not CmdCancel.Enabled
    CmdSalir.Enabled = Not CmdSalir.Enabled
    CmdEjecutar.Enabled = Not CmdEjecutar.Enabled
End Sub

Private Sub CmdEjecutar_Click()
    Dim Rpta, A As Integer
    Rpta = MsgBox("Esta seguro de asignar por defecto la opcion seleccionada", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
    If Rpta = vbNo Then
        Exit Sub
    End If
    
    'actualizamos todos los flag a 0
    For A = 1 To fg2.Rows - 1
        xCon.Execute "UPDATE con_formatostipo SET con_formatostipo.defecto = 0 WHERE (((con_formatostipo.idformato)=" & NulosN(LblIdLibro.Caption) & ") " _
            & " AND ((con_formatostipo.id)=" & NulosN(fg2.TextMatrix(A, 3)) & "))"
    Next A
    
    If xIdFormatos = 1 Then
        'para la tabla con_formatos
        xCon.Execute "UPDATE con_formatostipo SET con_formatostipo.defecto = -1 WHERE (((con_formatostipo.idformato)=" & NulosN(LblIdLibro.Caption) & ") " _
            & " AND ((con_formatostipo.id)=" & NulosN(fg2.TextMatrix(fg2.Row, 3)) & "))"
    Else
        'para la tabla con_analisis
        xCon.Execute "UPDATE con_formatostipo SET con_formatostipo.defecto = -1 WHERE (((con_formatostipo.idformato)=" & NulosN(LblIdLibro.Caption) & ") " _
            & " AND ((con_formatostipo.id)=" & NulosN(fg2.TextMatrix(fg2.Row, 3)) & "))"
    End If
    
    CargarFormato
End Sub

Private Sub CmdGrabar_Click()
    Dim Rst As New ADODB.Recordset
    Dim RstDet As New ADODB.Recordset
    Dim A, B As Integer
    
On Error GoTo LaCague
    xCon.BeginTrans

    If xIdFormatos = 1 Then
        xCon.Execute "DELETE * FROM con_formatostipodet  WHERE idformato = " & NulosN(LblIdLibro.Caption) & " AND idformatotipo = " & NulosN(fg2.TextMatrix(fg2.Row, 3)) & " "
        RST_Busq RstDet, "SELECT * FROM con_formatostipodet", xCon
    Else
    
    End If
    
    For B = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("idformato") = NulosN(LblIdLibro.Caption)
        RstDet("idformatotipo") = NulosN(fg2.TextMatrix(fg2.Row, 3))
        RstDet("id") = B
        RstDet("titulo") = NulosC(Fg1.TextMatrix(B, 10))
        RstDet("descripcion") = NulosC(Fg1.TextMatrix(B, 1))
        RstDet("abrev") = NulosC(Fg1.TextMatrix(B, 2))
        
        If NulosN(Fg1.TextMatrix(B, 4)) = -1 Then
            RstDet("mostrar") = -1
        Else
            RstDet("mostrar") = 0
        End If
        RstDet("orden") = NulosN(Fg1.TextMatrix(B, 3))
        RstDet("ancho") = NulosN(Fg1.TextMatrix(B, 5))
        RstDet("alineacion") = NulosN(Fg1.TextMatrix(B, 6))
        RstDet("nomcampo") = NulosC(Fg1.TextMatrix(B, 11))
        
        If NulosN(Fg1.TextMatrix(B, 7)) = -1 Then
            RstDet("imprimir") = -1
        Else
            RstDet("imprimir") = 0
        End If
        RstDet("anchoprin") = NulosN(Fg1.TextMatrix(B, 8))
        
        If NulosN(Fg1.TextMatrix(B, 9)) = -1 Then
            RstDet("totalizar") = -1
        Else
            RstDet("totalizar") = 0
        End If
        
        RstDet.Update
    Next B
    
    Set RstDet = Nothing
    MsgBox "El formato se guardo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    xCon.CommitTrans
    
    CargarFormato
    CmdCancel_Click
    Exit Sub
    
LaCague:
    MsgBox "No se puede guardar el registro por el siguiente motivo : " & NulosC(Err.Description), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    xCon.RollbackTrans
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 6 Then
        If NulosN(Fg1.TextMatrix(Fg1.Row, 6)) <> 1 And NulosN(Fg1.TextMatrix(Fg1.Row, 6)) <> 4 And NulosN(Fg1.TextMatrix(Fg1.Row, 6)) <> 7 Then
            Fg1.TextMatrix(Fg1.Row, 6) = ""
        End If
    End If
End Sub

Private Sub Fg1_RowColChange()
    LblDesc.Caption = Fg1.TextMatrix(Fg1.Row, 1)
End Sub

Private Sub fg2_RowColChange()
    CargarDetalle NulosN(LblIdLibro.Caption), NulosN(fg2.TextMatrix(fg2.Row, 3))
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        CargarFormato
    End If
End Sub

Sub CargarFormato()
    Dim Rst2 As New ADODB.Recordset
    Dim Rst1 As New ADODB.Recordset
    Dim A As Integer
    Dim xId As Integer

    If xIdFormatos = 1 Then
        Tabla1 = "con_formatos"
        Tabla2 = "con_formatostipo"
        Tabla3 = "con_formatostipodet"
    Else
        Tabla1 = "con_analisis"
        Tabla2 = "con_analisistipo"
        Tabla3 = "con_analisistipodet"
    End If
    
    If xIdFormatos = 1 Then
        RST_Busq Rst1, "SELECT * FROM con_formatos WHERE id = " & NulosN(LblIdLibro.Caption) & "", xCon
    Else
        RST_Busq Rst1, "SELECT * FROM con_analisis WHERE id = " & NulosN(LblIdLibro.Caption) & "", xCon
    End If
    
    If Rst1.RecordCount <> 0 Then
        TxtLibro.Text = Rst1("descripcion")
    End If
    Set Rst1 = Nothing
        
    If xIdFormatos = 1 Then
        RST_Busq Rst2, "SELECT con_formatostipo.idformato, con_formatostipo.descripcion, con_formatostipo.defecto, con_formatostipo.id" _
            & " From con_formatostipo WHERE (((con_formatostipo.idformato)= " & NulosN(LblIdLibro.Caption) & "))", xCon
    Else
        RST_Busq Rst2, "SELECT con_analisistipo.idformato, con_analisistipo.descripcion, con_analisistipo.defecto, con_analisistipo.id" _
            & " From con_analisistipo WHERE (((con_analisistipo.idformato)= " & NulosN(LblIdLibro.Caption) & "))", xCon
    End If

    If Rst2.RecordCount <> 0 Then
        fg2.Rows = 1
        Rst2.MoveFirst
        xId = Rst2("id") 'obtenemos el primer para mostrar su detalle
        
        For A = 1 To Rst2.RecordCount
            fg2.Rows = fg2.Rows + 1
            fg2.TextMatrix(A, 1) = Rst2("descripcion")
            If Rst2("defecto") = -1 Then
                fg2.TextMatrix(A, 2) = -1
            Else
                fg2.TextMatrix(A, 2) = 0
            End If
            
            fg2.TextMatrix(A, 3) = Rst2("id")
            
            Rst2.MoveNext
            If Rst2.EOF = True Then Exit For
        Next A
    End If
    
    CargarDetalle NulosN(LblIdLibro.Caption), xId
End Sub

Sub CargarDetalle(Libro As Integer, IdFormatoTipo As Integer)
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    
    LblDesc.Caption = ""
    If xIdFormatos = 1 Then
        RST_Busq Rst, "SELECT con_formatostipodet.* From con_formatostipodet WHERE (((con_formatostipodet.idformato)=" & Libro & ") AND ((con_formatostipodet.idformatotipo)=" & IdFormatoTipo & ")) ORDER BY orden", xCon
    Else
        RST_Busq Rst, "SELECT con_analisistipodet.idformato, con_analisistipodet.idformatotipo, con_analisistipodet.id, con_analisistipodet.descripcion, " _
            & " con_analisistipodet.abrev, con_analisistipodet.mostrar, con_analisistipodet.orden From con_analisistipodet " _
            & " WHERE (((con_analisistipodet.idformato)=" & Libro & ") AND ((con_analisistipodet.idformatotipo)=" & IdFormatoTipo & ")) ORDER BY orden", xCon
    End If
    
    Fg1.Rows = 1
    'Fg1.Editable = flexEDKbdMouse
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            
            Fg1.TextMatrix(A, 1) = NulosC(Rst("descripcion"))
            Fg1.TextMatrix(A, 2) = NulosC(Rst("abrev"))
            Fg1.TextMatrix(A, 3) = NulosN(Rst("orden"))
            
            If Rst("mostrar") = -1 Then
                Fg1.TextMatrix(A, 4) = -1
            Else
                Fg1.TextMatrix(A, 4) = 0
            End If
            
            Fg1.TextMatrix(A, 5) = NulosN(Rst("ancho"))
            Fg1.TextMatrix(A, 6) = NulosN(Rst("alineacion"))
            
            If Rst("imprimir") = -1 Then
                Fg1.TextMatrix(A, 7) = -1
            Else
                Fg1.TextMatrix(A, 7) = 0
            End If
            
            Fg1.TextMatrix(A, 8) = NulosN(Rst("anchoprin"))
            
            
            If Rst("totalizar") = -1 Then
                Fg1.TextMatrix(A, 9) = -1
            Else
                Fg1.TextMatrix(A, 9) = 0
            End If
            Fg1.TextMatrix(A, 10) = NulosC(Rst("titulo"))
            Fg1.TextMatrix(A, 11) = NulosC(Rst("nomcampo"))
            Fg1.TextMatrix(A, 12) = Rst("id")
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
        
        LblDesc.Caption = Fg1.TextMatrix(1, 1)
    End If
    Set Rst = Nothing
End Sub

Private Sub Form_Load()
    Dim Ruta As String
    
    SeEjecuto = False
    Fg1.ColWidth(10) = 0
    Fg1.ColWidth(11) = 0
    Fg1.ColWidth(12) = 0
    fg2.ColWidth(3) = 0
    Fg1.Editable = flexEDNone
    fg2.Editable = flexEDNone
    Ruta = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")
    
    Me.ScaleMode = 3
    CmdGrabar.Caption = ""
    CmdCancel.Caption = ""
    CmdEdit.Caption = ""
    CmdSalir.Caption = ""
    CmdEdit.Picture = LeerIcono(Ruta + "toolbar\2.ico", T32x32, Me, Me.BackColor)
    CmdGrabar.Picture = LeerIcono(Ruta + "toolbar\5.ico", T32x32, Me, Me.BackColor)
    CmdCancel.Picture = LeerIcono(Ruta + "toolbar\4.ico", T32x32, Me, Me.BackColor)
    CmdSalir.Picture = LeerIcono(Ruta + "toolbar\16.ico", T32x32, Me, Me.BackColor)
    CmdEjecutar.Picture = LeerIcono(Ruta + "toolbar\19.ico", T32x32, Me, Me.BackColor)
End Sub

