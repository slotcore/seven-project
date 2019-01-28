VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmSelCentroCosto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Seleccion de Centro de Costos"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   3870
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   9585
      _cx             =   16907
      _cy             =   6826
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
      FormatString    =   $"FrmSeleCentroCosto.frx":0000
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
   Begin VSFlex7Ctl.VSFlexGrid Fg2 
      Height          =   1470
      Left            =   30
      TabIndex        =   6
      Top             =   4680
      Width           =   9585
      _cx             =   16907
      _cy             =   2593
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmSeleCentroCosto.frx":00AD
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
      TabIndex        =   7
      Top             =   3900
      Width           =   9600
      Begin VB.CommandButton CmdAceptar 
         Height          =   450
         Left            =   6990
         Picture         =   "FrmSeleCentroCosto.frx":0126
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Width           =   1170
      End
      Begin VB.CommandButton CmdSalir 
         Height          =   450
         Left            =   8205
         Picture         =   "FrmSeleCentroCosto.frx":24AC
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   210
         Width           =   1170
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
         TabIndex        =   8
         Text            =   "TxtNivAct"
         Top             =   210
         Width           =   645
      End
      Begin VB.CommandButton CmdAddSel 
         Height          =   450
         Left            =   4350
         Picture         =   "FrmSeleCentroCosto.frx":42FE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   210
         Width           =   2190
      End
      Begin VB.CommandButton CmdBajaNiv 
         Height          =   450
         Left            =   2400
         Picture         =   "FrmSeleCentroCosto.frx":82D8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   195
         Width           =   1485
      End
      Begin VB.CommandButton CmdSubNiv 
         Height          =   450
         Left            =   180
         Picture         =   "FrmSeleCentroCosto.frx":AD4E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   195
         Width           =   1485
      End
   End
End
Attribute VB_Name = "FrmSelCentroCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rst As New ADODB.Recordset
Dim Nivel As Integer
Dim NumNivel As Integer
Dim Mostrando As Boolean
Dim SeEjecuto As Boolean
Public RstSele As New ADODB.Recordset

Private Sub CmdAceptar_Click()
    Dim a As Integer
    PreparaRst
    For a = 1 To Fg2.Rows - 1
        RstSele.AddNew
        RstSele("idcencos") = Fg2.TextMatrix(a, 3)
        RstSele("codigo") = Fg2.TextMatrix(a, 1)
        RstSele("descripcion") = Fg2.TextMatrix(a, 2)
    Next a
    CmdSalir_Click
End Sub

Private Sub CmdAddSel_Click()
    Dim a, B As Integer
    
    'revisamos que no dupliquen centros de costos
    For a = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(a, 3)) = -1 Then
            For B = 1 To Fg2.Rows - 1
                If Fg1.TextMatrix(a, 1) = Fg2.TextMatrix(B, 1) Then
                    Fg1.TextMatrix(a, 3) = False
                End If
            Next B
        End If
    Next a
    
    'copiamos los centros de costo
    For a = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(a, 3)) = -1 Then
            Fg2.Rows = Fg2.Rows + 1
            Fg2.TextMatrix(Fg2.Rows - 1, 1) = Fg1.TextMatrix(a, 1)
            Fg2.TextMatrix(Fg2.Rows - 1, 2) = Fg1.TextMatrix(a, 2)
            Fg2.TextMatrix(Fg2.Rows - 1, 3) = Fg1.TextMatrix(a, 4)
        End If
    Next a
End Sub

Private Sub CmdBajaNiv_Click()
    If Val(Fg1.TextMatrix(Fg1.Row, 5)) = 0 Then
        MsgBox "El centro de costo seleccionado no contiene sub centros de costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    MuestraNivel 1
End Sub

Private Sub CmdSalir_Click()
    Set Rst = Nothing
    Unload Me
End Sub

Private Sub CmdSubNiv_Click()
     If NumNivel = 1 Then
        MsgBox "Esta en el nivel superior", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
     End If
     MuestraNivel 2
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Mostrando = True Then Exit Sub
    If Fg1.Col = 3 Then
        If Fg1.TextMatrix(Fg1.Row, 5) = 1 Then
            Fg1.TextMatrix(Fg1.Row, 3) = False
        End If
    End If
End Sub

Private Sub Fg1_DblClick()
    If Fg1.TextMatrix(Fg1.Row, 5) = 0 Then
        MsgBox "El centro de costo seleccionado no tiene sub centros de costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    MuestraNivel 1
End Sub
        
Sub MuestraNivel(Accion As Integer)
    'Accion = 1  Incrementa el nivel
    'Accion = 2  Disminuir el nivel
    
    Dim a As Integer
    Dim xCad As String
    
    If Accion = 1 Then
        Nivel = Nivel + 2
        NumNivel = NumNivel + 1
        TxtNivAct.Text = Trim(Str(NumNivel))
        
        RST_Busq Rst, "SELECT con_centrocosto.id, con_centrocosto.codigo, con_centrocosto.descripcion, Len(Trim([codigo])) AS [long], " _
            & " con_centrocosto.tipo, Mid([codigo],1,2) AS xcod From con_centrocosto " _
            & " WHERE (((Len(Trim([codigo])))=" & Nivel & " ) AND ((Mid([codigo],1," & Nivel - 2 & "))='" & Fg1.TextMatrix(Fg1.Row, 1) & "'))", xCon
    Else
        If Nivel = 2 Then Exit Sub
        Nivel = Nivel - 2
        NumNivel = NumNivel - 1
        TxtNivAct.Text = Trim(Str((NumNivel)))
        
        xCad = mID(Fg1.TextMatrix(Fg1.Row, 1), 1, Nivel - 2)
        RST_Busq Rst, "SELECT con_centrocosto.id, con_centrocosto.codigo, con_centrocosto.descripcion, Len(Trim([codigo])) AS [long], " _
            & " con_centrocosto.tipo, Mid([codigo],1,2) AS xcod " _
            & " From con_centrocosto WHERE (((Len(Trim([codigo])))=" & Nivel & ") AND " _
            & " ((Mid([codigo],1," & Nivel - 2 & ")) = '" & xCad & "'))", xCon
    End If

    If Rst.RecordCount = 0 Then
        If Accion = 1 Then Nivel = Nivel - 2
        MsgBox "El centro de costos seleccionado no tiene sub centros de costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Fg1.Rows = 1
    Mostrando = True
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For a = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(a, 1) = Rst("codigo")
            Fg1.TextMatrix(a, 2) = Rst("descripcion")
            Fg1.TextMatrix(a, 4) = Rst("id")
            Fg1.TextMatrix(a, 5) = Rst("tipo")
            
            If Rst("tipo") = 1 Then
                With Fg1
                    .Select a, 1, a, 3
                    .FillStyle = flexFillRepeat
                    .CellForeColor = &H800000
                End With
            End If
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next a
    End If
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
            MsgBox "El centro de costo seleccionado no tiene sub centros de costo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
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
    
    If KeyCode = 45 Then  'tecla insert
        If Fg1.TextMatrix(Fg1.Row, 5) = 1 Then Exit Sub
        
        Dim B As Integer
        If Fg1.TextMatrix(Fg1.Row, 4) = "1" Then Exit Sub
        
        For B = 1 To Fg2.Rows - 1
            If Fg1.TextMatrix(Fg1.Row, 1) = Fg2.TextMatrix(B, 1) Then
                MsgBox "El centro de costos seleccionado ya fue agregado a la lista", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            End If
        Next B
        
        'copiamos los centros de costo
        Fg2.Rows = Fg2.Rows + 1
        Fg2.TextMatrix(Fg2.Rows - 1, 1) = Fg1.TextMatrix(Fg1.Row, 1)
        Fg2.TextMatrix(Fg2.Rows - 1, 2) = Fg1.TextMatrix(Fg1.Row, 2)
        Fg2.TextMatrix(Fg2.Rows - 1, 3) = Fg1.TextMatrix(Fg1.Row, 4)
     
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        Dim a As Integer
        Nivel = 2
        SeEjecuto = True
        NumNivel = 1
        TxtNivAct.Text = Trim(Str(NumNivel))
        
        RST_Busq Rst, "SELECT con_centrocosto.id, con_centrocosto.codigo, con_centrocosto.descripcion, Len(Trim([codigo])) AS [long], tipo " _
            & " From con_centrocosto WHERE (((Len(Trim([codigo])))=2))", xCon
    
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            For a = 1 To Rst.RecordCount
                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(a, 1) = Rst("codigo")
                Fg1.TextMatrix(a, 2) = Rst("descripcion")
                Fg1.TextMatrix(a, 4) = Rst("id")
                Fg1.TextMatrix(a, 5) = Rst("tipo")
                
                If Rst("tipo") = 1 Then
                    With Fg1
                        .Select a, 1, a, 3
                        .FillStyle = flexFillRepeat
                        .CellForeColor = &H800000
                    End With
                End If
                
                Rst.MoveNext
                If Rst.EOF = True Then
                    Exit For
                End If
            Next a
        End If
    End If
End Sub

Private Sub Form_Load()
    Fg1.Rows = 1
    Fg1.ColWidth(4) = 0
    Fg1.ColWidth(5) = 0
    
    Fg2.ColWidth(3) = 0
    Fg2.Rows = 1
    SeEjecuto = False
End Sub

Sub PreparaRst()
    'Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(3, 3) As String

    xCampos(0, 0) = "idcencos":     xCampos(0, 1) = "N":      xCampos(0, 2) = "3"
    xCampos(1, 0) = "codigo":       xCampos(1, 1) = "C":      xCampos(1, 2) = "12"
    xCampos(2, 0) = "descripcion":  xCampos(2, 1) = "C":      xCampos(2, 2) = "150"
    
    Set RstSele = xFun.CrearRstTMP(xCampos)
    RstSele.Open
End Sub
