VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Begin VB.Form FrmSelecciona 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccion de Registros"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      Caption         =   "( Opciones de Seleccion )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   765
      Left            =   60
      TabIndex        =   3
      Top             =   15
      Width           =   10875
      Begin VB.OptionButton Option2 
         BackColor       =   &H00808000&
         Caption         =   "Deseleccionar todo"
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
         Height          =   210
         Left            =   2775
         TabIndex        =   5
         Top             =   360
         Width           =   2040
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "Seleccionar todo"
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
         Height          =   210
         Left            =   390
         TabIndex        =   4
         Top             =   360
         Width           =   2040
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Height          =   765
      Left            =   45
      TabIndex        =   0
      Top             =   5205
      Width           =   10890
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   405
         Left            =   5460
         TabIndex        =   2
         Top             =   255
         Width           =   1140
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   405
         Left            =   4290
         TabIndex        =   1
         Top             =   255
         Width           =   1140
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   4350
      Left            =   45
      TabIndex        =   6
      Top             =   855
      Width           =   10890
      _cx             =   19209
      _cy             =   7673
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
      BackColor       =   14417405
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   128
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14417405
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
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmSelecciona.frx":0000
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
Attribute VB_Name = "FrmSelecciona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Acepto As Boolean
Public Rst As New ADODB.Recordset

Private Sub CmdAceptar_Click()
    Acepto = True
    Me.Hide
    
    Dim A As Integer
    Dim xCampInd As String
    Dim xNumColInd As Integer
    
    For A = LBound(xCampos) To UBound(xCampos)
        If xCampos(A, 4) = "S" Then
            xCampInd = xCampos(A, 1)
            xNumColInd = A
        End If
        
        If A = UBound(xCampos) - 1 Then
            Exit For
        End If
    Next A
    
    Set Rst.ActiveConnection = Nothing
    For A = 1 To Fg1.Rows - 1
        If F_NulosN(Fg1.TextMatrix(A, 0)) = 0 Then
            Rst.Filter = adFilterNone
            If xCampos(xNumColInd, 3) = "C" Then
                Rst.Filter = "" & xCampInd & " = '" & Fg1.TextMatrix(A, xNumColInd + 1) & "'"
            Else
                Rst.Filter = "" & xCampInd & " = " & Fg1.TextMatrix(A, xNumColInd + 1) & ""
            End If
            If Rst.RecordCount <> 0 Then
                Rst.Delete
            End If
        End If
    Next A
    Rst.Filter = adFilterNone
    Me.Hide
End Sub

Private Sub CmdCancelar_Click()
    Acepto = False
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim A As Integer
    
    F_RST_Busq Rst, xSQLCad, xConeccion
    
    For A = LBound(xCampos) To UBound(xCampos)
        Fg1.Cols = Fg1.Cols + 1
        Fg1.TextMatrix(0, Fg1.Cols - 1) = xCampos(A, 0)
        Fg1.ColWidth(Fg1.Cols - 1) = xCampos(A, 2)
        
        If A = UBound(xCampos) - 1 Then
            Exit For
        End If
    Next A
    
    Dim B As Integer
    Dim xCol As Integer
    If Rst.RecordCount = 0 Then
        MsgBox "No se han encontrado registros con las condiciones especificadas", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Acepto = False
        Unload Me
        Exit Sub
    End If
    
    Rst.MoveFirst
    For A = 1 To Rst.RecordCount
        Fg1.Rows = Fg1.Rows + 1
        
        xCol = 1
        For B = LBound(xCampos) To UBound(xCampos)
            If UCase(xCampos(B, 3)) = "C" Then
                Fg1.TextMatrix(A, xCol) = F_NulosC(Rst(xCampos(B, 1)))
            Else
                Fg1.TextMatrix(A, xCol) = Format(F_NulosN(Rst(xCampos(B, 1))), "0.00")
            End If
            
            If B = UBound(xCampos) - 1 Then
                Exit For
            End If
            xCol = xCol + 1
        Next B
        Rst.MoveNext
        If Rst.EOF = True Then
            Exit For
        End If
    Next A
End Sub

Private Sub Form_Load()
    Fg1.Cols = 1
    Fg1.Rows = 1
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.Editable = flexEDKbdMouse
End Sub

Private Sub Option1_Click()
    Dim A As Integer
    For A = 1 To Fg1.Rows - 1
        Fg1.TextMatrix(A, 0) = -1
    Next A
End Sub

Private Sub Option2_Click()
    Dim A As Integer
    For A = 1 To Fg1.Rows - 1
        Fg1.TextMatrix(A, 0) = 0
    Next A
End Sub
