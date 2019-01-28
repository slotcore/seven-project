VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmVerMovimiento 
   Caption         =   "Procesos - OPeraciones Sobre el Registro"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   2400
      Left            =   0
      TabIndex        =   2
      Top             =   30
      Width           =   7365
      _cx             =   12991
      _cy             =   4233
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
      BackColor       =   12320247
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   64
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   12320247
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
      Rows            =   50
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmVerMovimiento.frx":0000
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
   Begin VB.TextBox TxtSQL 
      Height          =   285
      Left            =   7560
      TabIndex        =   1
      Text            =   "TxtSQL"
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   2460
      Width           =   7365
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   315
         Left            =   2760
         TabIndex        =   0
         Top             =   240
         Width           =   1230
      End
   End
End
Attribute VB_Name = "FrmVerMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rst As New ADODB.Recordset
Dim QueHace  As Integer
Dim SeEjecuto As Double

Private Sub CmdAceptar_Click()
    Unload Me
    Set Rst = Nothing
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        Dim A As Integer
        
        SeEjecuto = True
        RST_Busq Rst, NulosC(TxtSQL.Text), xCon
        If Rst.RecordCount <> 0 Then
            Rst.MoveFirst
            Fg1.Rows = 1
            For A = 1 To Rst.RecordCount
                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(Rst("formulario"))
                Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(Rst("apenom"))
                Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(Rst("operacion"))
                Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(NulosC(Rst("fchope")), "dd/mm/yy")
                Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(NulosC(Rst("horini")), "hh:mm:ss")
                Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(NulosC(Rst("horfin")), "hh:mm:ss")
                
                Rst.MoveNext
                If Rst.EOF = True Then Exit For
                
            Next A
        Else
            MsgBox "No se ha registrado movimientos sobre el registro especificado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Set Rst = Nothing
            Unload Me
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    Fg1.Rows = 1
    Fg1.ColWidth(1) = 0
    
    Fg1.AutoSearch = flexSearchFromTop
    Fg1.ExplorerBar = flexExSortShow
    
End Sub

