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
   Begin VB.CommandButton Command20 
      Caption         =   "Reporte de Stocks"
      Height          =   735
      Left            =   510
      TabIndex        =   18
      Top             =   4410
      Width           =   1275
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Honorarios"
      Height          =   735
      Left            =   570
      TabIndex        =   17
      Top             =   2190
      Width           =   1275
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Costo de Personal"
      Height          =   735
      Left            =   7620
      TabIndex        =   16
      Top             =   3720
      Width           =   1275
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Linea"
      Height          =   735
      Left            =   6330
      TabIndex        =   15
      Top             =   2070
      Width           =   1275
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Asistencia"
      Height          =   735
      Left            =   7710
      TabIndex        =   14
      Top             =   2760
      Width           =   1275
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Grupos"
      Height          =   735
      Left            =   6210
      TabIndex        =   13
      Top             =   4590
      Width           =   1275
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Maestro de costo"
      Height          =   735
      Left            =   7710
      TabIndex        =   12
      Top             =   1950
      Width           =   1275
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Registro de tareas"
      Height          =   735
      Left            =   7710
      TabIndex        =   11
      Top             =   1080
      Width           =   1275
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Maestro de Almacen"
      Height          =   735
      Left            =   480
      TabIndex        =   10
      Top             =   3510
      Width           =   1275
   End
   Begin VB.CommandButton RepCosto 
      Caption         =   "Reporte de costo"
      Height          =   735
      Left            =   6210
      TabIndex        =   9
      Top             =   3780
      Width           =   1275
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Registro de costo personal"
      Height          =   735
      Left            =   6240
      TabIndex        =   8
      Top             =   2940
      Width           =   1275
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Consulta Produccion"
      Height          =   735
      Left            =   7680
      TabIndex        =   7
      Top             =   4530
      Width           =   1275
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Orden de Compra"
      Height          =   735
      Left            =   2370
      TabIndex        =   6
      Top             =   990
      Width           =   1275
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Registro Produccion"
      Height          =   735
      Left            =   6270
      TabIndex        =   5
      Top             =   5400
      Width           =   1275
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cotizacion"
      Height          =   735
      Left            =   2340
      TabIndex        =   4
      Top             =   1770
      Width           =   1275
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Consulta de Compras"
      Height          =   735
      Left            =   540
      TabIndex        =   3
      Top             =   1020
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Registro Produccion V2"
      Height          =   735
      Left            =   7710
      TabIndex        =   2
      Top             =   5340
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
      Caption         =   "Compras"
      Height          =   735
      Left            =   540
      TabIndex        =   0
      Top             =   240
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ' EJECUTA MENU
    Dim xform As New sgi2_compras.Compras
    xform.IdMenu = 218
    xform.Idusuario = 1
    xform.RegCompras2 xCon, AP_MESTRA, 0
    Set xform = Nothing
End Sub

Private Sub Command19_Click()
    Dim xfrm As New sgi2_compras.Compras
    xfrm.IdMenu = 8
    xfrm.Idusuario = 1
    xfrm.RegHonorarios xCon, CInt(Mid(Date, 4, 2)), 0
    Set xfrm = Nothing
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    Dim xfrm As New sgi2_compras.Compras
    xfrm.IdMenu = 7
    xfrm.Idusuario = 1
    xfrm.RepCompras xCon
    Set xfrm = Nothing
End Sub

Private Sub Command5_Click()
    Dim xfrm As New sgi2_compras.Compras
    xfrm.IdMenu = 7
    xfrm.Idusuario = 1
    xfrm.ManCotizacionCompra xCon, CInt(Mid(Date, 4, 2))
End Sub

Private Sub Command7_Click()
    Dim xfrm As New sgi2_compras.Compras
    xfrm.IdMenu = 7
    xfrm.Idusuario = 1
    xfrm.OrdenCompra xCon
End Sub

Private Sub Form_Load()
    Me.KeyPreview = True
    Main
End Sub
