VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmSeteaMes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema - Mes de trabajo"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3090
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   3090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOk 
      Height          =   630
      Left            =   735
      Picture         =   "FrmSeteaMes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2145
      Width           =   750
   End
   Begin VB.CommandButton CmdSalir 
      Height          =   630
      Left            =   1545
      Picture         =   "FrmSeteaMes.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2145
      Width           =   750
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg3 
      Height          =   2070
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   3015
      _cx             =   5318
      _cy             =   3651
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
      FormatString    =   $"FrmSeteaMes.frx":0614
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
End
Attribute VB_Name = "FrmSeteaMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rst As New ADODB.Recordset
Dim SeEjecuto As Boolean

Private Sub CmdOk_Click()
    If Fg3.Row < 1 Then Exit Sub
    xMes = Val(Fg3.TextMatrix(Fg3.Row, 3))
    MsgBox "Ha seleccionado el mes de trabajo " + Fg3.TextMatrix(Fg3.Row, 1), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    CmdSalir_Click
End Sub

Private Sub CmdSalir_Click()
    Set Rst = Nothing
    Unload Me
End Sub


Private Sub Fg3_DblClick()
    If Fg3.Row < 0 Or Fg3.Col < 0 Then Exit Sub
    CmdOk_Click
End Sub

Private Sub Fg3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then CmdOk_Click
End Sub

Private Sub Form_Activate()
    If SeEjecuto = True Then Exit Sub
   
    SeEjecuto = False
    
    Dim a As Integer
    RST_Busq Rst, "SELECT * FROM con_meses ORDER BY id", xCon
    
    Fg3.Rows = 1
    Rst.MoveFirst
    For a = 1 To Rst.RecordCount
        Fg3.Rows = Fg3.Rows + 1
        Fg3.TextMatrix(a, 1) = Rst("descripcion")
        Fg3.TextMatrix(a, 3) = Rst("id")
        If Rst("cerrado") = -1 Then
            Fg3.TextMatrix(a, 2) = -1
        Else
            Fg3.TextMatrix(a, 2) = 0
        End If
        Rst.MoveNext
        
        If Rst.EOF = True Then
            Exit For
        End If
    Next a
    
    If Fg3.Rows > 1 Then
        Fg3.Col = 1
        Fg3.Row = 1
        Fg3.SetFocus
    Else
        CmdOk.SetFocus
    End If
    Set Rst = Nothing
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Set Rst = Nothing
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    Fg3.ColWidth(3) = 0
End Sub
