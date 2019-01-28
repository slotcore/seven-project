VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form ProgressForm 
   Caption         =   "Cargando"
   ClientHeight    =   1275
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProcessProgressBar 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label SubProcessLabel 
      AutoSize        =   -1  'True
      Caption         =   "SubProcessLabel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1000
      Width           =   1155
   End
   Begin VB.Label SubDescriptionLabel 
      Caption         =   "SubDescriptionLabel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   1005
      Width           =   6735
   End
   Begin VB.Label DescriptionLabel 
      Caption         =   "DescriptionLabel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   8055
   End
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mProceso As String
Private mSubProceso As String
Private mDescripcion As String
Private mSubDescripcion As String
Private mMin As Long
Private mMax As Long
Private mValue As Long

Public Property Get Proceso() As String
    Proceso = mProceso
End Property
Public Property Let Proceso(ByVal NewValue As String)
    mProceso = NewValue
    Me.Caption = mProceso
End Property

Public Property Get SubProceso() As String
    SubProceso = mSubProceso
End Property
Public Property Let SubProceso(ByVal NewValue As String)
    mSubProceso = NewValue
    SubProcessLabel.Caption = mSubProceso
End Property

Public Property Get Descripcion() As String
    Descripcion = mDescripcion
End Property
Public Property Let Descripcion(ByVal NewValue As String)
    mDescripcion = NewValue
    DescriptionLabel.Caption = mDescripcion
End Property

Public Property Get SubDescripcion() As String
    SubDescripcion = mSubDescripcion
End Property
Public Property Let SubDescripcion(ByVal NewValue As String)
    mSubDescripcion = NewValue
    SubDescriptionLabel.Caption = mSubDescripcion
End Property

Public Property Get Min() As Long
    Min = mMin
End Property
Public Property Let Min(ByVal NewValue As Long)
    mMin = NewValue
    ProcessProgressBar.Min = mMin
End Property

Public Property Get Max() As Long
    Max = mMax
End Property
Public Property Let Max(ByVal NewValue As Long)
    mMax = NewValue
    ProcessProgressBar.Max = mMax
End Property

Public Property Get Value() As Long
    Value = mValue
End Property
Public Property Let Value(ByVal NewValue As Long)
    mValue = NewValue
    ProcessProgressBar.Value = NewValue
End Property

Private Sub Form_Load()
    Me.ZOrder 0
    SubProcessLabel.Caption = ""
    SubDescriptionLabel.Caption = ""
End Sub
