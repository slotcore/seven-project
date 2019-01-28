VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command21 
      Caption         =   "Conectar"
      Height          =   495
      Left            =   4680
      TabIndex        =   24
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Desencriptar"
      Height          =   495
      Left            =   4680
      TabIndex        =   23
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   2760
      TabIndex        =   22
      Text            =   "Text3"
      Top             =   6360
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   2760
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Encriptar"
      Height          =   495
      Left            =   4680
      TabIndex        =   20
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   2760
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Reporte de Stocks"
      Height          =   735
      Left            =   540
      TabIndex        =   18
      Top             =   5040
      Width           =   1275
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Ingreso/Salida"
      Height          =   735
      Left            =   570
      TabIndex        =   17
      Top             =   1920
      Width           =   1275
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Costo de Personal"
      Height          =   735
      Left            =   7380
      TabIndex        =   16
      Top             =   3000
      Width           =   1275
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Toma Inventario"
      Height          =   735
      Left            =   2520
      TabIndex        =   15
      Top             =   1080
      Width           =   1395
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Asistencia"
      Height          =   735
      Left            =   7470
      TabIndex        =   14
      Top             =   2040
      Width           =   1275
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Movimiento Automatico"
      Height          =   735
      Left            =   8880
      TabIndex        =   13
      Top             =   600
      Width           =   1275
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Despacho Automatico"
      Height          =   735
      Left            =   2520
      TabIndex        =   12
      Top             =   3600
      Width           =   1275
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Almacenaje Automatico"
      Height          =   735
      Left            =   2520
      TabIndex        =   11
      Top             =   2760
      Width           =   1275
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Maestro de Almacen"
      Height          =   735
      Left            =   570
      TabIndex        =   10
      Top             =   2760
      Width           =   1275
   End
   Begin VB.CommandButton RepCosto 
      Caption         =   "Reporte de costo"
      Height          =   735
      Left            =   5760
      TabIndex        =   9
      Top             =   1080
      Width           =   1275
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Transferencias"
      Height          =   735
      Left            =   2520
      TabIndex        =   8
      Top             =   1920
      Width           =   1395
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Consulta Produccion"
      Height          =   735
      Left            =   7440
      TabIndex        =   7
      Top             =   3810
      Width           =   1275
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Devolucion"
      Height          =   735
      Left            =   600
      TabIndex        =   6
      Top             =   3810
      Width           =   1275
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Registro Produccion"
      Height          =   735
      Left            =   5910
      TabIndex        =   5
      Top             =   2760
      Width           =   1275
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Consulta de Items"
      Height          =   735
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   1275
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Items"
      Height          =   735
      Left            =   570
      TabIndex        =   3
      Top             =   180
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Registro Produccion V2"
      Height          =   735
      Left            =   7470
      TabIndex        =   2
      Top             =   4620
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   735
      Left            =   9630
      TabIndex        =   1
      Top             =   6120
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recepcion"
      Height          =   735
      Left            =   540
      TabIndex        =   0
      Top             =   1050
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim xfrm As New SGI2_almacen.Almacen
    xfrm.IdMenu = 253
    xfrm.Idusuario = 1
    xfrm.RecepcionAlmacen xCon, CInt(Mid(Date, 4, 2))
    Set xfrm = Nothing
End Sub

Private Sub Command10_Click()
    Dim xfrm As New SGI2_almacen.Almacen
    xfrm.IdMenu = 88
    xfrm.Idusuario = 1
    xfrm.ManAlmacen xCon
    Set xfrm = Nothing
End Sub

Private Sub Command11_Click()
    Dim xfrm As New SGI2_almacen.Almacen
    xfrm.IdMenu = 8
    xfrm.Idusuario = 1
    xfrm.ManAlmacenajeAuto xCon
    Set xfrm = Nothing
End Sub

Private Sub Command12_Click()
    Dim xfrm As New SGI2_almacen.Almacen
    xfrm.IdMenu = 8
    xfrm.Idusuario = 1
    xfrm.ManDespachoAuto xCon
    Set xfrm = Nothing
End Sub

Private Sub Command15_Click()
    Dim xfrm As New SGI2_almacen.Almacen
    xfrm.IdMenu = 8
    xfrm.Idusuario = 1
    xfrm.TomaInventario xCon
    Set xfrm = Nothing
End Sub

Private Sub Command17_Click()
'    Dim F As New Funciones
'    Text2.Text = F.Encriptar(Text1.Text)
'    Set F = Nothing
End Sub

Private Sub Command18_Click()
'    Dim F As New Funciones
'    Text3.Text = F.Desencriptar(Text2.Text)
'    Set F = Nothing
End Sub

Private Sub Command21_Click()
'    Dim data As New SistemaData.Database
'    MsgBox (data.Password)
'    Set data = Nothing
End Sub

Private Sub Command19_Click()
    Dim xfrm As New SGI2_almacen.Almacen
    xfrm.IdMenu = 8
    xfrm.Idusuario = 1
    xfrm.IngresoAlmacen xCon, CInt(Mid(Date, 4, 2))
    Set xfrm = Nothing
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command20_Click()
    Dim xfrm As New SGI2_almacen.Almacen
    xfrm.IdMenu = 88
    xfrm.Idusuario = 1
    xfrm.ConsultaStock xCon
    Set xfrm = Nothing
End Sub

Private Sub Command4_Click()
    Dim xfrm As New SGI2_almacen.Almacen
    xfrm.IdMenu = 7
    xfrm.Idusuario = 1
    xfrm.MantItem xCon, 1
    Set xfrm = Nothing
End Sub

Private Sub Command5_Click()
    Dim xfrm As New SGI2_almacen.Almacen
    xfrm.IdMenu = 7
    xfrm.Idusuario = 1
    xfrm.ConsultaItems xCon
    Set xfrm = Nothing
End Sub

Private Sub Command7_Click()
    Dim xfrm As New SGI2_almacen.Almacen
    xfrm.IdMenu = 253
    xfrm.Idusuario = 1
    xfrm.DevolucionAlmacen xCon, CInt(Mid(Date, 4, 2))
    Set xfrm = Nothing
End Sub

Private Sub Command9_Click()
    Dim xfrm As New SGI2_almacen.Almacen
    xfrm.IdMenu = 8
    xfrm.Idusuario = 1
    xfrm.TransferenciaAlmacen xCon, CInt(Mid(Date, 4, 2))
    Set xfrm = Nothing
End Sub

Private Sub Form_Load()
    Me.KeyPreview = True
    Main
End Sub
