VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmSelCentroCosto2 
   Caption         =   "Centro de Costos"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8685
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   5325
      Left            =   45
      TabIndex        =   0
      Top             =   390
      Width           =   8595
      _cx             =   15161
      _cy             =   9393
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
      BackColor       =   14614269
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12582912
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14614269
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
      FormatString    =   $"FrmSelCentroCosto2.frx":0000
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
      Height          =   840
      Left            =   45
      TabIndex        =   2
      Top             =   5670
      Width           =   8625
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   450
         Left            =   4350
         TabIndex        =   4
         Top             =   240
         Width           =   1230
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   450
         Left            =   3075
         TabIndex        =   3
         Top             =   240
         Width           =   1230
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Centros de Costos"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   30
      TabIndex        =   1
      Top             =   45
      Width           =   8640
   End
End
Attribute VB_Name = "FrmSelCentroCosto2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rst As New ADODB.Recordset
Dim SeEjecuto  As Boolean
Public RstDat As New ADODB.Recordset
Public EnviarRST As Boolean

Private Sub CmdAceptar_Click()
    Dim A As Integer
    
    RST_Busq RstDat, "SELECT con_centrocosto.id, con_centrocosto.codigo, con_centrocosto.descripcion, con_centrocosto.tipo " _
        & " From con_centrocosto WHERE (((con_centrocosto.id)=99999))", xCon
    
    RstDat.ActiveConnection = Nothing
    
    For A = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(A, 3)) = -1 Then
            RstDat.AddNew
            RstDat("id") = Fg1.TextMatrix(A, 4)
            RstDat("codigo") = Fg1.TextMatrix(A, 1)
            RstDat("descripcion") = Fg1.TextMatrix(A, 2)
            If Fg1.TextMatrix(A, 5) = 1 Then
                RstDat("tipo") = 1
            Else
                RstDat("tipo") = 0
            End If
        End If
    Next A
    EnviarRST = True
    Me.Hide
End Sub

Private Sub CmdCancelar_Click()
    EnviarRST = False
    Me.Hide
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim A, valor As Integer
    If Col = 3 Then
        If NulosN(Fg1.TextMatrix(Fg1.Row, 5)) = 0 Then
            'Fg1.TextMatrix(Fg1.Row, 3) = 0
            Exit Sub
        Else
            Dim xAmchoCodigo As Integer
            Dim CodPadre As String
            
            CodPadre = Trim(Fg1.TextMatrix(Fg1.Row, 1))
            xAmchoCodigo = Len(Trim(Fg1.TextMatrix(Fg1.Row, 1)))
            
            If NulosN(Fg1.TextMatrix(Fg1.Row, 3)) = -1 Then
                valor = 1
            Else
                valor = 0
            End If
            
            For A = Fg1.Row To Fg1.Rows - 1
                If mID(Fg1.TextMatrix(A, 1), 1, xAmchoCodigo) = CodPadre Then
                    Fg1.TextMatrix(A, 3) = valor
                Else
                    Fg1.TextMatrix(A, 3) = 0
                End If
            Next A
        End If
    End If
End Sub

Private Sub Fg1_EnterCell()
    If Fg1.Col = 3 Then
        If NulosN(Fg1.TextMatrix(Fg1.Row, 5)) = 0 Then
            Fg1.TextMatrix(Fg1.Row, 3) = 0
            Exit Sub
        End If
        
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        Dim A As Integer
        SeEjecuto = True
        RST_Busq Rst, "SELECT con_centrocosto.* From con_centrocosto ORDER BY con_centrocosto.codigo", xCon
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            Fg1.Rows = 1
            For A = 1 To Rst.RecordCount
                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(Fg1.Rows - 1, 4) = Rst("id")
                Fg1.TextMatrix(Fg1.Rows - 1, 5) = Rst("tipo")
                If Rst("tipo") = 1 Then
                    FlexFormatoCelda Fg1, Fg1.Rows - 1, 1, &H800000, True, &HDEFEFD, Rst("codigo")
                    FlexFormatoCelda Fg1, Fg1.Rows - 1, 2, &H800000, True, &HDEFEFD, Rst("descripcion")
                Else
                    FlexFormatoCelda Fg1, Fg1.Rows - 1, 1, &H80000012, False, &HDEFEFD, Rst("codigo")
                    FlexFormatoCelda Fg1, Fg1.Rows - 1, 2, &H80000012, False, &HDEFEFD, Rst("descripcion")
                End If
                
                Rst.MoveNext
                If Rst.EOF = True Then Exit For
            Next A
        End If
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    Fg1.ColWidth(4) = 0
    Fg1.ColWidth(5) = 0
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.BackColorSel = &H80&
End Sub
