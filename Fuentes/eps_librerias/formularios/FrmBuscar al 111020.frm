VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmBuscar 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00808000&
      Caption         =   "[ Criterio de Busqueda]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   30
      TabIndex        =   10
      Top             =   0
      Width           =   5340
      Begin VB.TextBox TxtCriterio 
         BackColor       =   &H00C0C000&
         Height          =   300
         Left            =   75
         TabIndex        =   0
         Text            =   "TxtCriterio"
         Top             =   255
         Width           =   5175
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808000&
      Caption         =   "[ Campo de Busqueda]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   5400
      TabIndex        =   9
      Top             =   0
      Width           =   3375
      Begin VB.ComboBox CboCampos 
         BackColor       =   &H00C0C000&
         Height          =   315
         Left            =   225
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   240
         Width           =   2940
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4155
      Left            =   15
      TabIndex        =   1
      Top             =   705
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   7329
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12632064
      HeadLines       =   1.5
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      Height          =   765
      Left            =   5100
      TabIndex        =   8
      Top             =   4800
      Width           =   3660
      Begin VB.CommandButton CmdCan 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   1830
         TabIndex        =   6
         Top             =   240
         Width           =   1350
      End
      Begin VB.CommandButton CmdAce 
         Caption         =   "&Aceptar"
         Height          =   390
         Left            =   450
         TabIndex        =   5
         Top             =   240
         Width           =   1350
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Height          =   765
      Left            =   45
      TabIndex        =   7
      Top             =   4800
      Width           =   5040
      Begin VB.OptionButton Option2 
         BackColor       =   &H00808000&
         Caption         =   "&Cualquier Parte del Campo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Left            =   2370
         TabIndex        =   3
         Top             =   330
         Width           =   2610
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "&Principio del Campo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Left            =   135
         TabIndex        =   2
         Top             =   315
         Width           =   2220
      End
   End
End
Attribute VB_Name = "FrmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RstBusca As New ADODB.Recordset
Dim xreg As Integer
Dim Salir As Boolean
Dim CaracteresNumericos As String
Dim CaracteresAlfaNumericos As String

Private Sub CboCampos_Click()
    xCampoBusca = BuscaCampoLista(Trim(CboCampos.Text), 0, 1, xCampos)
    If BuscaCampoLista(Trim(CboCampos.Text), 0, 3, xCampos) = "M" Then
        CboCampos.Text = ""
        TxtCriterio.Text = ""
        MsgBox "No se puede realizar la busqueda sobre un campo MEMO, seleccione otro campo", vbInformation + vbOKOnly + vbDefaultButton1, "Busqueda"
        Exit Sub
    End If
    TxtCriterio.Text = ""
    xOrdenado = xCampoBusca
    TxtCriterio.SetFocus
    RstBusca.Sort = xOrdenado
End Sub

Private Sub CmdAce_Click()
    Cancelado = False
    Me.Hide
End Sub

Private Sub CmdCan_Click()
    Set RstBusca = Nothing
    Cancelado = True
    Unload Me
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdAce_Click
    End If
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 96 Then KeyCode = 48  'tecla Nº 0
'    If KeyCode = 97 Then KeyCode = 49  'tecla Nº 1
'    If KeyCode = 98 Then KeyCode = 50  'tecla Nº 2
'    If KeyCode = 99 Then KeyCode = 51  'tecla Nº 3
'    If KeyCode = 100 Then KeyCode = 52  'tecla Nº 4
'    If KeyCode = 101 Then KeyCode = 53  'tecla Nº 5
'    If KeyCode = 102 Then KeyCode = 54  'tecla Nº 6
'    If KeyCode = 103 Then KeyCode = 58  'tecla Nº 7
'    If KeyCode = 104 Then KeyCode = 56  'tecla Nº 8
'    If KeyCode = 105 Then KeyCode = 57  'tecla Nº 9
'
    If KeyCode <> 13 And KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 9 And KeyCode <> 34 And KeyCode <> 33 Then
        TxtCriterio = ""
        TxtCriterio.Text = Chr(KeyCode)
        TxtCriterio.SetFocus
        'SendKeys vbKeyEnd
    End If
End Sub

Private Sub Form_Activate()
    
On Error GoTo LaCague
    xTitulo = "Busqueda de Registros"
    Salir = False
    TxtCriterio.Text = ""
    
    CrearLista
    LLenarCombo
    'CboCampos = xCampoBusca
    CboCampos = BuscaCampoLista(xCampoBusca, 1, 0, xCampos)
    
    If EjecutaSQL = True Then
        F_RST_Busq RstBusca, xSQLCad, xConeccion
    Else
        Set RstBusca = xRstConsulta
    End If
    
    If RstBusca.State <> 0 Then
        If RstBusca.RecordCount = 0 Then
            Set RstBusca = Nothing
            MsgBox "No se han encontrado registros para la busqueda actual", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Cancelado = True
            Unload Me
            Exit Sub
        End If
        
        RstBusca.Sort = xOrdenado
        DataGrid1.HeadFont.Bold = True
        DataGrid1.HoldFields
        DataGrid1.MarqueeStyle = dbgHighlightRow
        Set DataGrid1.DataSource = RstBusca
        If xFormaBusca = Principio Then
            Option1.Value = True
        Else
            Option2.Value = True
        End If
        
        TxtCriterio.SetFocus
        
        If xCriterio = "" Then
            TxtCriterio = ""
            TxtCriterio.SetFocus
        Else
    '        TxtCriterio = xCriterio
    '        RstBusca.MoveFirst
    '        RstBusca.Find "" & xCampoBusca & " LIKE '" & Trim(TxtCriterio.Text) & "*'"
    '        DataGrid1.SetFocus
    '        If RstBusca.EOF = True Then
    '            TxtCriterio.Text = Left(Trim(TxtCriterio.Text), Len(Trim(TxtCriterio.Text)) - 1)
    '            RstBusca.Bookmark = xreg
    '            SendKeys "{END}"
    '        End If
        End If
    End If
    Exit Sub

LaCague:
    MsgBox "No se pudo cargar la busqueda por el siguiente motivo : " & Err.Description, vbInformation + vbOKOnly + vbDefaultButton1
    Exit Sub
End Sub

Sub LLenarCombo()
    Dim A As Integer
    For A = LBound(xCampos) To UBound(xCampos)
        CboCampos.AddItem xCampos(A, 0)  'muestra los titulos de los campos en el menu
        If A = UBound(xCampos) - 1 Then
            Exit For
        End If
    Next A
End Sub

Sub CrearLista()
    Dim A As Integer
    Dim B As Integer
    Dim C As Integer
    B = 1
    
    For A = LBound(xCampos) To UBound(xCampos)
        
        DataGrid1.Columns.Item(A).Caption = xCampos(A, 0)
        DataGrid1.Columns.Item(A).DataField = xCampos(A, 1)
        DataGrid1.Columns.Item(A).Width = xCampos(A, 2)
        If xCampos(A, 3) = "N" Then
            'DataGrid1.Columns.Item(A).NumberFormat = "0.00"
            DataGrid1.Columns.Item(A).Alignment = dbgRight
        End If
        If xCampos(A, 3) = "C" Then
            DataGrid1.Columns.Item(A).Alignment = dbgLeft
        End If
        If xCampos(A, 3) = "D" Then
            DataGrid1.Columns.Item(A).NumberFormat = "dd/mm/yy"
            DataGrid1.Columns.Item(A).Alignment = dbgCenter
        End If
        If A = UBound(xCampos) - 1 Then
            Exit For
        End If
        
        DataGrid1.Columns.Add (B)
        B = B + 1
    Next A
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        CmdCan_Click
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = xTitulo
    Cancelado = True
    CaracteresNumericos = "0123456789.-'%&$()!¡¿?" & Chr(8)
    CaracteresAlfaNumericos = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZabcdefghijklmnñopqrstuvwxyz01234567890 " & Chr(8)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If Cancelado <> False Then Cancelado = True
End Sub

Private Sub Option1_Click()
    RstBusca.Filter = adFilterNone
    RstBusca.MoveFirst
End Sub

Private Sub Option2_Click()
    RstBusca.Filter = adFilterNone
    RstBusca.MoveFirst
End Sub

Private Sub TxtCriterio_Change()
    If Salir = True Then Exit Sub
    If F_NulosC(TxtCriterio.Text) = "" Then
        If RstBusca.State = 1 Then
            If RstBusca.RecordCount <> 0 Then RstBusca.MoveFirst: Exit Sub
        End If
    End If
    
    If CboCampos.Text = "" Then
        MsgBox "No ha especificado en campo donde se efectuara a busqueda", vbInformation + vbOKOnly + vbDefaultButton1, "Busqueda"
        CboCampos.SetFocus
        Exit Sub
    End If
    If F_NulosC(TxtCriterio.Text) = "" Then Exit Sub
    If Mid(F_NulosC(TxtCriterio.Text), Len(F_NulosC(TxtCriterio.Text)), 1) = "'" Or Mid(F_NulosC(TxtCriterio.Text), Len(F_NulosC(TxtCriterio.Text)), 1) = "%" Then
        TxtCriterio.Text = ""
        Exit Sub
    End If
    
    If Option1.Value = True Then
        If F_NulosC(TxtCriterio.Text) <> "" Then
            RstBusca.MoveFirst
            If BuscaCampoLista(F_NulosC(CboCampos.Text), 0, 3, xCampos) = "C" Then
                RstBusca.Find "" & xCampoBusca & " LIKE '" & F_NulosC(TxtCriterio.Text) & "*'"
            End If
            If BuscaCampoLista(Trim(CboCampos.Text), 0, 3, xCampos) = "N" Then
                RstBusca.Find "" & xCampoBusca & " = " & Trim(TxtCriterio.Text) & ""
            End If
        
            If RstBusca.EOF = True Then
                TxtCriterio.Text = Left(Trim(TxtCriterio.Text), Len(Trim(TxtCriterio.Text)) - 1)
                TxtCriterio.SetFocus
                SendKeys "{END}"
            End If
            DataGrid1.SetFocus
            TxtCriterio.SetFocus
            'SendKeys "{END}"
        Else
            If RstBusca.State = 1 Then
                RstBusca.MoveFirst
            End If
        End If
    End If
End Sub

Sub MostrarFiltro()
    If Salir = True Then TxtCriterio.Text = "": Exit Sub
    If CboCampos.Text = "" Then
        MsgBox "No ha especificado en campo donde se efectuara a busqueda", vbInformation + vbOKOnly + vbDefaultButton1, "Busqueda"
        CboCampos.SetFocus
        Exit Sub
    End If
    
    If Trim(TxtCriterio.Text) <> "" Then
        'RstBusca.MoveFirst
        If BuscaCampoLista(Trim(CboCampos.Text), 0, 3, xCampos) = "C" Then
            'RstBusca.Find "" & xCampoBusca & " LIKE '*" & Trim(TxtCriterio.Text) & "*'"
            RstBusca.Filter = "" & xCampoBusca & " LIKE '*" & Trim(TxtCriterio.Text) & "*'"
        End If
        If BuscaCampoLista(Trim(CboCampos.Text), 0, 3, xCampos) = "N" Then
            'RstBusca.Find "" & xCampoBusca & " = " & Trim(TxtCriterio.Text) & ""
            RstBusca.Filter = "" & xCampoBusca & " = " & Trim(TxtCriterio.Text) & ""
        End If
    
        If RstBusca.EOF = True Then
            TxtCriterio.Text = Left(Trim(TxtCriterio.Text), Len(Trim(TxtCriterio.Text)) - 1)
            TxtCriterio.SetFocus
            SendKeys "{END}"
        End If
    Else
        If RstBusca.State = 1 Then
            RstBusca.MoveFirst
        End If
    End If
    TxtCriterio.Text = ""
End Sub

Private Sub TxtCriterio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then Salir = True: Exit Sub
    Salir = False
    
    If Option1.Value = True Then
        If KeyAscii = 13 And Len(Trim(TxtCriterio.Text)) > 0 Then
            DataGrid1.SetFocus
            'SendKeys vbTab
        Else
            If BuscaCampoLista(F_NulosC(CboCampos.Text), 0, 3, xCampos) = "N" Then
                If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
            End If
            If BuscaCampoLista(F_NulosC(CboCampos.Text), 0, 3, xCampos) = "C" Then
                If InStr(CaracteresAlfaNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
            End If
        End If
    Else
        If BuscaCampoLista(F_NulosC(CboCampos.Text), 0, 3, xCampos) = "N" Then
            If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then
                If KeyAscii <> 13 Then KeyAscii = 0
            End If
        End If
        
        If KeyAscii = 13 And Len(Trim(TxtCriterio.Text)) > 0 Then
            If Option1.Value = True Then
                Exit Sub
            Else
                MostrarFiltro
                SendKeys vbTab
                Exit Sub
            End If
        Else
            RstBusca.Filter = adFilterNone
        End If
        If KeyAscii = 9 Then
            DataGrid1.SetFocus
            Exit Sub
        End If
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtCriterio_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        DataGrid1.SetFocus
    End If
    If KeyCode = 40 Then
        DataGrid1.SetFocus
        SendKeys "{DOWN}"
    End If
    If KeyCode = 219 Then
        TxtCriterio.Text = ""
    End If
End Sub
