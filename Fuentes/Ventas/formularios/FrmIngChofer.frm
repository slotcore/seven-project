VERSION 5.00
Begin VB.Form FrmIngChofer 
   Caption         =   "Ingreso de Choferes"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2595
      Left            =   30
      TabIndex        =   9
      Top             =   -60
      Width           =   4890
      Begin VB.TextBox TxtCategoria 
         Height          =   300
         Left            =   1605
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "TxtCategoria"
         Top             =   1425
         Width           =   1320
      End
      Begin VB.TextBox TxtNumBreve 
         Height          =   300
         Left            =   1605
         MaxLength       =   15
         TabIndex        =   3
         Text            =   "TxtNumBreve"
         Top             =   1125
         Width           =   1320
      End
      Begin VB.TextBox TxtNumPlaca 
         Height          =   300
         Left            =   1605
         MaxLength       =   20
         TabIndex        =   6
         Text            =   "TxtNumPlaca"
         Top             =   2160
         Width           =   1320
      End
      Begin VB.TextBox TxtVehiculo 
         Height          =   300
         Left            =   1605
         MaxLength       =   20
         TabIndex        =   5
         Text            =   "TxtVehiculo"
         Top             =   1860
         Width           =   2415
      End
      Begin VB.TextBox TxtNombre 
         Height          =   300
         Left            =   1605
         MaxLength       =   40
         TabIndex        =   2
         Text            =   "TxtNombre"
         Top             =   825
         Width           =   2415
      End
      Begin VB.TextBox TxtApeMat 
         Height          =   300
         Left            =   1605
         MaxLength       =   40
         TabIndex        =   1
         Text            =   "TxtApeMat"
         Top             =   525
         Width           =   2415
      End
      Begin VB.TextBox TxtApePat 
         Height          =   300
         Left            =   1605
         MaxLength       =   40
         TabIndex        =   0
         Text            =   "TxtApePat"
         Top             =   225
         Width           =   2415
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Categoria"
         Height          =   195
         Left            =   270
         TabIndex        =   17
         Top             =   1470
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nº Brevete"
         Height          =   195
         Left            =   270
         TabIndex        =   16
         Top             =   1170
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nº Placa"
         Height          =   195
         Left            =   270
         TabIndex        =   14
         Top             =   2205
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Marca Vehiculo"
         Height          =   195
         Left            =   270
         TabIndex        =   13
         Top             =   1890
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nombres"
         Height          =   195
         Left            =   270
         TabIndex        =   12
         Top             =   885
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Materno"
         Height          =   195
         Left            =   270
         TabIndex        =   11
         Top             =   585
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Apellido Paterno"
         Height          =   195
         Left            =   270
         TabIndex        =   10
         Top             =   285
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Height          =   930
      Left            =   30
      TabIndex        =   15
      Top             =   2505
      Width           =   4890
      Begin VB.CommandButton CmdSalir 
         Caption         =   "CmdSalir"
         Height          =   720
         Left            =   2460
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   150
         Width           =   810
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "CmdGrabar"
         Height          =   720
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   150
         Width           =   810
      End
   End
End
Attribute VB_Name = "FrmIngChofer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SeEjecuto As Boolean

Sub Blanquea()
    TxtApePat.Text = ""
    TxtApeMat.Text = ""
    TxtNombre.Text = ""
    TxtVehiculo.Text = ""
    TxtNumPlaca.Text = ""
    TxtNumBreve.Text = ""
    TxtCategoria.Text = ""
End Sub

Private Sub CmdGrabar_Click()
    'Grabar
    If Grabar = True Then
        CmdSalir_Click
    End If
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        Blanquea
        TxtApePat.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim Ruta As String
    
    SeEjecuto = False
    Ruta = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABM", "RUTAS")
    
    Me.ScaleMode = 3
    CmdGrabar.Caption = ""
    CmdSalir.Caption = ""
    
    On Error GoTo LaCague
    
    CmdGrabar.Picture = LeerIcono(Ruta + "toolbar\5.ico", T32x32, Me, Me.BackColor)
    CmdSalir.Picture = LeerIcono(Ruta + "toolbar\16.ico", T32x32, Me, Me.BackColor)
    Exit Sub
    
LaCague:
    CmdGrabar.Caption = "Grabar"
    CmdSalir.Caption = "Salir"
End Sub

Private Sub TxtApeMat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtApePat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtCategoria_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumBreve_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumPlaca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtVehiculo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Function Grabar() As Boolean
'Modificado 20/01/11 Johan Castro
'           Cambiar tipo de datos de xId, xIdCho, xIdVehi As Double antes Integer
'           Agregar linea de codigo para registrar el historial de empleados, vehiculos y chofer

    Grabar = False
    If NulosC(TxtApePat.Text) = "" Then
        MsgBox "No ha especificado el apellido Paterno", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtApePat.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtApeMat.Text) = "" Then
        MsgBox "No ha especificado el apellido Materno", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtApeMat.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtNombre.Text) = "" Then
        MsgBox "No ha especificado los nombres del chofer", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNombre.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtNumBreve.Text) = "" Then
        MsgBox "No ha especificado el numero de brevete", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumBreve.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtCategoria.Text) = "" Then
        MsgBox "No ha especificado la categoria del chofer", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtCategoria.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtVehiculo.Text) = "" Then
        MsgBox "No ha especificado la marca del vehiculo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtVehiculo.SetFocus
        Exit Function
    End If
    
    If TxtNumPlaca.Text = "" Then
        MsgBox "No ha especificado el numero de placa del vehiculo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumPlaca.SetFocus
        Exit Function
    End If
    
    Grabar = True
    Dim xCon2 As New ADODB.Connection
    Set xCon2 = AbriConeccion2(xCon)
    Dim Rst As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    Dim Rst3 As New ADODB.Recordset
    Dim xId As Double
    Dim xIdCho As Double
    Dim xIdVehi As Double
    
    On Error GoTo LaCague
    xCon.BeginTrans
    
    RST_Busq Rst, "SELECT * FROM pla_empleados", xCon2
    
    xId = HallaCodigoTabla("pla_empleados", xCon2, "id")
    
    Rst.AddNew
    Rst("id") = xId
    Rst("apepat") = TxtApePat.Text
    Rst("apemat") = TxtApeMat.Text
    Rst("nom") = TxtNombre.Text
    Rst("idsex") = 1
    Rst("idtipdoc") = 1
    Rst("idnac") = 193
    Rst.Update
    
    'grabamos el movimiento en la tabla var_edicion - empleados
    GrabarOperacion xIdUsuario, 59, 1, Time, Time, Date, xCon, xId

    
    RST_Busq Rst3, "SELECT * FROM mae_vehiculo", xCon
    xIdVehi = HallaCodigoTabla("mae_vehiculo", xCon, "id")
    Rst3.AddNew
    Rst3("id") = xIdVehi
    Rst3("marca") = NulosC(TxtVehiculo.Text)
    Rst3("numpla") = NulosC(TxtNumPlaca.Text)
    Rst3.Update
    
    'grabamos el movimiento en la tabla var_edicion - unidades de transporte
    GrabarOperacion xIdUsuario, 112, 1, Time, Time, Date, xCon, xIdVehi
    
    RST_Busq Rst2, "SELECT * FROM mae_chofer", xCon
    xIdCho = HallaCodigoTabla("mae_chofer", xCon, "id")
    Rst2.AddNew
    Rst2("id") = xIdCho
    Rst2("idvehiculo") = xIdVehi
    Rst2("idPer") = xId
    Rst2("numbre") = NulosC(TxtNumBreve.Text)
    Rst2("categoria") = NulosC(TxtCategoria.Text)
    Rst2.Update
    
    'grabamos el movimiento en la tabla var_edicion - choferes
    GrabarOperacion xIdUsuario, 111, 1, Time, Time, Date, xCon, xIdCho
    
    xCon.CommitTrans
    MsgBox "El nuevo chofer se guardo con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Grabar = True
    Exit Function

LaCague:
    MsgBox "No se ha podido guardar el registro por el siguiente motivo " & Err.Description
    xCon.RollbackTrans
    Grabar = False
End Function
