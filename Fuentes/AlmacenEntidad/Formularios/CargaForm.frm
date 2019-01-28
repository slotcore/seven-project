VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form CargaForm 
   Caption         =   "Procesando"
   ClientHeight    =   1155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1155
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar LoadProgressBar 
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   529
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label AccionLabel 
      AutoSize        =   -1  'True
      Caption         =   "AccionLabel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1005
   End
   Begin VB.Label DetalleLabel 
      Alignment       =   2  'Center
      Caption         =   "DetalleLabel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   4425
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   950
      Index           =   0
      Left            =   120
      Top             =   100
      Width           =   5925
   End
End
Attribute VB_Name = "CargaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
