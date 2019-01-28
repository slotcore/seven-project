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
      Caption         =   "Reporte de Linea"
      Height          =   735
      Left            =   450
      TabIndex        =   18
      Top             =   4470
      Width           =   1275
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Levantar Pedidos"
      Height          =   735
      Left            =   480
      TabIndex        =   17
      Top             =   1980
      Width           =   1275
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Costo de Personal"
      Height          =   735
      Left            =   5340
      TabIndex        =   16
      Top             =   3660
      Width           =   1275
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Linea"
      Height          =   735
      Left            =   4050
      TabIndex        =   15
      Top             =   2010
      Width           =   1275
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Asistencia"
      Height          =   735
      Left            =   5430
      TabIndex        =   14
      Top             =   2700
      Width           =   1275
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Grupos"
      Height          =   735
      Left            =   3930
      TabIndex        =   13
      Top             =   4530
      Width           =   1275
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Maestro de costo"
      Height          =   735
      Left            =   5430
      TabIndex        =   12
      Top             =   1890
      Width           =   1275
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Registro de tareas"
      Height          =   735
      Left            =   5430
      TabIndex        =   11
      Top             =   1020
      Width           =   1275
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Maestro de tareas"
      Height          =   735
      Left            =   5370
      TabIndex        =   10
      Top             =   180
      Width           =   1275
   End
   Begin VB.CommandButton RepCosto 
      Caption         =   "Reporte de costo"
      Height          =   735
      Left            =   3930
      TabIndex        =   9
      Top             =   3720
      Width           =   1275
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Registro de costo personal"
      Height          =   735
      Left            =   3960
      TabIndex        =   8
      Top             =   2880
      Width           =   1275
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Consulta Produccion"
      Height          =   735
      Left            =   2490
      TabIndex        =   7
      Top             =   4500
      Width           =   1275
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Receta"
      Height          =   735
      Left            =   4050
      TabIndex        =   6
      Top             =   180
      Width           =   1275
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Registro Produccion"
      Height          =   735
      Left            =   600
      TabIndex        =   5
      Top             =   3450
      Width           =   1275
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Planeacion de Produccion"
      Height          =   735
      Left            =   4140
      TabIndex        =   4
      Top             =   1110
      Width           =   1275
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ventas"
      Height          =   735
      Left            =   570
      TabIndex        =   3
      Top             =   180
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Registro Produccion V2"
      Height          =   735
      Left            =   2010
      TabIndex        =   2
      Top             =   3450
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   735
      Left            =   5430
      TabIndex        =   1
      Top             =   4560
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guia de remision"
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
    ' EJECUTA MENU
    Dim xfrm As New sgi2_ventas.ventas
    xfrm.IdMenu = 17
    xfrm.Idusuario = 1
    
    xfrm.GuiasRemision xCon, CInt(Mid(Date, 4, 2))
    Set xfrm = Nothing
End Sub

Private Sub Command19_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_ventas.ventas
    xfrm.IdMenu = 17
    xfrm.Idusuario = 1
    
    xfrm.LevantarPedidos xCon
    Set xfrm = Nothing
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    ' EJECUTA MENU
    Dim xfrm As New sgi2_ventas.ventas
    xfrm.IdMenu = 18
    xfrm.Idusuario = 1
    xfrm.ventas xCon, CInt(Mid(Date, 4, 2))
    Set xfrm = Nothing
End Sub

Private Sub Form_Load()
    Me.KeyPreview = True
    Main
End Sub
