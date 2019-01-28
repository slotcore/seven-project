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
   Begin VB.CommandButton Command18 
      Caption         =   "Cronograma de Produccion V2"
      Height          =   735
      Left            =   8010
      TabIndex        =   20
      Top             =   2190
      Width           =   1275
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Orden de Producion v2"
      Height          =   735
      Left            =   8070
      TabIndex        =   19
      Top             =   180
      Width           =   1275
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Reporte de Linea"
      Height          =   735
      Left            =   450
      TabIndex        =   18
      Top             =   4470
      Width           =   1275
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Reporte de Planeacion"
      Height          =   735
      Left            =   570
      TabIndex        =   17
      Top             =   2310
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
      Caption         =   "Reporte de Produccion"
      Height          =   735
      Left            =   3840
      TabIndex        =   14
      Top             =   4800
      Width           =   1275
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Solic. Materi v2"
      Height          =   735
      Left            =   9600
      TabIndex        =   13
      Top             =   210
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
      Left            =   480
      TabIndex        =   7
      Top             =   5280
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
      Left            =   540
      TabIndex        =   5
      Top             =   3390
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
      Caption         =   "Crono Prod"
      Height          =   735
      Left            =   570
      TabIndex        =   3
      Top             =   180
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Registro Produccion V2"
      Height          =   735
      Left            =   8070
      TabIndex        =   2
      Top             =   1170
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   735
      Left            =   9870
      TabIndex        =   1
      Top             =   6270
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Solic. Materi"
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
'    Dim xFrm As New sgi2_produccion.produccion
'    xFrm.IdMenu = 54
'    xFrm.Idusuario = 1
'    xFrm.GenOrdenProduccion xCon, CInt(Mid(Date, 4, 2))
'    Set xFrm = Nothing
End Sub

Private Sub Command10_Click()
     Dim xFrm As New sgi2_produccion.produccion
    xFrm.ManTareas xCon
End Sub

Private Sub Command11_Click()
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.IngresoTareas xCon, CInt(Mid(Date, 4, 2))
    Set xFrm = Nothing
End Sub

Private Sub Command12_Click()
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.ConfigurarCosto xCon
    Set xFrm = Nothing
End Sub

Private Sub Command13_Click()
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.IdMenu = 54
    xFrm.Idusuario = 1
    xFrm.GenSolicitudMat xCon, CInt(Mid(Date, 4, 2))
    Set xFrm = Nothing
End Sub

Private Sub Command14_Click()
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.RepCompProduccion xCon
    Set xFrm = Nothing
End Sub

Private Sub Command15_Click()
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.CronogramaMantLinea xCon
    Set xFrm = Nothing
End Sub

Private Sub Command16_Click()
    Dim xFrm As New sgi2_produccion.produccion
    'xFrm.RepCosto xCon
    xFrm.CostoProduccion xCon
End Sub

Private Sub Command17_Click()
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.IdMenu = 52
    xFrm.Idusuario = 1
    xFrm.GenOrdenProd xCon, CInt(Mid(Date, 4, 2))
    Set xFrm = Nothing
End Sub

Private Sub Command18_Click()
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.IdMenu = 51
    xFrm.Idusuario = 1
    xFrm.CronogramaProduccion xCon
    Set xFrm = Nothing
End Sub

Private Sub Command19_Click()
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.RepPlaneacion xCon
    Set xFrm = Nothing
End Sub

'Private Sub Command10_Click()
'    Dim xfrm As New sgi2_produccion.Produccion
'    xfrm.RepCosto xCon
'End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command20_Click()
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.RepLinea xCon
    Set xFrm = Nothing
End Sub

Private Sub Command3_Click()
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.IdMenu = 92
    xFrm.Idusuario = 1
    xFrm.OrdenProduccion xCon, CInt(Mid(Date, 4, 2))
End Sub

Private Sub Command4_Click()
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.IdMenu = 51
    xFrm.Idusuario = 1
    xFrm.CronogramaProduccion xCon
    Set xFrm = Nothing
End Sub

Private Sub Command6_Click()
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.IdMenu = 92
    xFrm.Idusuario = 1
    xFrm.OrdenProduccion xCon, CInt(Mid(Date, 4, 2))
End Sub

Private Sub Command7_Click()
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.IdMenu = 92
    xFrm.Idusuario = 1
    xFrm.MamRecetas xCon
End Sub

Private Sub Command8_Click()
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.RepProduccion xCon
End Sub

Private Sub Command9_Click()
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.CostoProduccion xCon
End Sub

Private Sub Form_Load()
    Me.KeyPreview = True
    Main
End Sub

Private Sub RepCosto_Click()
    Dim xFrm As New sgi2_produccion.produccion
    xFrm.RepCosto xCon
End Sub
