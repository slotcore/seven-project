VERSION 5.00
Begin VB.MDIForm MDIMenu 
   BackColor       =   &H8000000C&
   Caption         =   "S.G.I. - Sistema de Gestion Integral V. 2.0"
   ClientHeight    =   5400
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8460
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu menu1 
      Caption         =   "Almacen"
      Begin VB.Menu Menu1_1 
         Caption         =   "Maestro Productos"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "Maestro de Insumos"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "Gestion de Guias"
      End
      Begin VB.Menu Menu1_4 
         Caption         =   "Salidas de Almacen"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "Ventas"
      Begin VB.Menu Menu2_1 
         Caption         =   "Proyeccion de Ventas"
      End
      Begin VB.Menu Menu2_2 
         Caption         =   "Ordenes de Compra"
      End
      Begin VB.Menu Menu2_3 
         Caption         =   "Plan de Ventas"
      End
   End
End
Attribute VB_Name = "MDImENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub menu1_3_Click()
    FrmSalidas.Show
End Sub

Private Sub Menu1_4_Click()
    FrmSalInsumos.Show
End Sub

Private Sub Menu2_1_Click()
    FrmPVEstimado.Show
End Sub

Private Sub Menu2_2_Click()
    FrmContratos.Show
End Sub
