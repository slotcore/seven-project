VERSION 5.00
Begin VB.Form FrmIngRapItems 
   Caption         =   "Ingreso Rapido de Items"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   900
      Left            =   105
      TabIndex        =   34
      Top             =   3195
      Width           =   6555
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancelar"
         Height          =   480
         Left            =   3300
         TabIndex        =   10
         Top             =   255
         Width           =   1065
      End
      Begin VB.CommandButton CmdAcepta 
         Caption         =   "&Aceptar"
         Height          =   480
         Left            =   2205
         TabIndex        =   9
         Top             =   255
         Width           =   1065
      End
   End
   Begin VB.CommandButton CmdBusMoneda 
      Height          =   240
      Left            =   2010
      Picture         =   "FrmIngRapItems.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2610
      Width           =   240
   End
   Begin VB.TextBox TxtDescripcion 
      Height          =   480
      Left            =   1365
      MaxLength       =   100
      TabIndex        =   5
      Text            =   "TxtDescripcion"
      Top             =   1710
      Width           =   5280
   End
   Begin VB.TextBox TxtCodPro 
      Height          =   300
      Left            =   1365
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   0
      Text            =   "TxtCodPro"
      Top             =   60
      Width           =   1770
   End
   Begin VB.CommandButton CmdBusClase 
      Height          =   240
      Left            =   2010
      Picture         =   "FrmIngRapItems.frx":0132
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1035
      Width           =   240
   End
   Begin VB.CommandButton CmdBusSubClase 
      Height          =   240
      Left            =   2010
      Picture         =   "FrmIngRapItems.frx":0264
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1350
      Width           =   240
   End
   Begin VB.CommandButton CmdBusUnidad 
      Height          =   240
      Left            =   2010
      Picture         =   "FrmIngRapItems.frx":0396
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2295
      Width           =   240
   End
   Begin VB.CommandButton CmdBusTipiTem 
      Height          =   240
      Left            =   2010
      Picture         =   "FrmIngRapItems.frx":04C8
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   405
      Width           =   240
   End
   Begin VB.CommandButton CmdBusFam 
      Height          =   240
      Left            =   2010
      Picture         =   "FrmIngRapItems.frx":05FA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   720
      Width           =   240
   End
   Begin VB.CommandButton CmdBusTipMovimiento 
      Height          =   240
      Left            =   2010
      Picture         =   "FrmIngRapItems.frx":072C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2925
      Width           =   240
   End
   Begin VB.TextBox TxtTipPro 
      Height          =   300
      Left            =   1365
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "TxtTipPro"
      Top             =   375
      Width           =   915
   End
   Begin VB.TextBox TxtIdSubClase 
      Height          =   300
      Left            =   1365
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "TxtIdSubClase"
      Top             =   1320
      Width           =   915
   End
   Begin VB.TextBox TxtIdClase 
      Height          =   300
      Left            =   1365
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "TxtIdClase"
      Top             =   1005
      Width           =   915
   End
   Begin VB.TextBox TxtIdFamilia 
      Height          =   300
      Left            =   1365
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "TxtIdFamilia"
      Top             =   690
      Width           =   915
   End
   Begin VB.TextBox TxtUnidad 
      Height          =   300
      Left            =   1365
      MaxLength       =   5
      TabIndex        =   6
      Text            =   "TxtUnidad"
      Top             =   2265
      Width           =   915
   End
   Begin VB.TextBox TxtIdMon 
      Height          =   300
      Left            =   1365
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "TxtIdMon"
      Top             =   2580
      Width           =   915
   End
   Begin VB.TextBox TxtIdTipmov 
      Height          =   300
      Left            =   1365
      MaxLength       =   1
      TabIndex        =   8
      Text            =   "TxtIdTipmov"
      Top             =   2895
      Width           =   915
   End
   Begin VB.Label LblPrefijo 
      Caption         =   "LblPrefijo"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   3390
      TabIndex        =   38
      Top             =   105
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label LblPrefijo1 
      Caption         =   "LblPrefijo1"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   4095
      TabIndex        =   37
      Top             =   105
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label LblPrefijo2 
      Caption         =   "LblPrefijo2"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   4875
      TabIndex        =   36
      Top             =   105
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label LblPrefijo3 
      Caption         =   "LblPrefijo3"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   5715
      TabIndex        =   35
      Top             =   105
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Unidad"
      Height          =   195
      Index           =   12
      Left            =   75
      TabIndex        =   33
      Top             =   2295
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Descripción"
      Height          =   195
      Index           =   6
      Left            =   75
      TabIndex        =   32
      Top             =   1770
      Width           =   840
   End
   Begin VB.Label LblClase 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblClase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   2325
      TabIndex        =   31
      Top             =   1005
      Width           =   4320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Clase"
      Height          =   195
      Index           =   1
      Left            =   75
      TabIndex        =   30
      Top             =   1035
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Index           =   7
      Left            =   75
      TabIndex        =   29
      Top             =   90
      Width           =   495
   End
   Begin VB.Label Moneda 
      AutoSize        =   -1  'True
      Caption         =   "Moneda"
      Height          =   195
      Left            =   75
      TabIndex        =   28
      Top             =   2640
      Width           =   585
   End
   Begin VB.Label LblSubClase 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblSubClase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   2325
      TabIndex        =   27
      Top             =   1320
      Width           =   4320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Sub Clase"
      Height          =   195
      Index           =   2
      Left            =   75
      TabIndex        =   26
      Top             =   1350
      Width           =   720
   End
   Begin VB.Label LblDescUnidad 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblDescUnidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   2325
      TabIndex        =   25
      Top             =   2265
      Width           =   4320
   End
   Begin VB.Label LblMoneda 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblMoneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   2325
      TabIndex        =   24
      Top             =   2580
      Width           =   4320
   End
   Begin VB.Label LblTipoPro 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblTipoPro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   2325
      TabIndex        =   23
      Top             =   375
      Width           =   4320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Item"
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   22
      Top             =   405
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Familia"
      Height          =   195
      Index           =   3
      Left            =   75
      TabIndex        =   21
      Top             =   720
      Width           =   480
   End
   Begin VB.Label LblFamilia 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblFamilia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   2325
      TabIndex        =   20
      Top             =   690
      Width           =   4320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Movimiento"
      Height          =   195
      Index           =   23
      Left            =   75
      TabIndex        =   19
      Top             =   2925
      Width           =   1170
   End
   Begin VB.Label LblTipoMovi 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LblTipoMovi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   2310
      TabIndex        =   18
      Top             =   2895
      Width           =   4320
   End
End
Attribute VB_Name = "FrmIngRapItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public xIdNewItem As Integer
Dim CaracteresNumericos As String
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim CODIGOTMP As String
Dim RstTem As New ADODB.Recordset
Public xIdProducto As Integer

Private Sub CmdAcepta_Click()
    Dim xCampos(10, 4) As String
    Dim xId As Integer
    Dim A As Integer
    
    TxtCodPro.Text = LblPrefijo.Caption & LblPrefijo1.Caption & LblPrefijo2.Caption & LblPrefijo3.Caption
    
On Error GoTo LaCague
    xCon.BeginTrans
    If QueHace = 1 Then
        xId = HallaCodigoTabla("alm_inventario", xCon, "id")
        xIdProducto = xId
    End If
    
    'Columna    | Descripcion
    '------------------------
    '0          | campo
    '1          | Valor
    '2          | requerido
    '3          | tipo
    '4          | msj error
    
    '--------------------------------
    'GRABAMOS LA CABECERA DE LA LETRA
    xCampos(0, 0) = "id":             xCampos(0, 1) = Str(xId):                   xCampos(0, 2) = "S":    xCampos(0, 3) = "N":    xCampos(0, 4) = "":
    xCampos(1, 0) = "codpro":         xCampos(1, 1) = TxtCodPro.Text:             xCampos(1, 2) = "S":    xCampos(1, 3) = "N":    xCampos(1, 4) = ""
    xCampos(2, 0) = "descripcion":    xCampos(2, 1) = TxtDescripcion.Text:        xCampos(2, 2) = "S":    xCampos(2, 3) = "C":    xCampos(2, 4) = ""
    xCampos(3, 0) = "desctec":        xCampos(3, 1) = TxtDescripcion.Text:        xCampos(3, 2) = "S":    xCampos(3, 3) = "C":    xCampos(3, 4) = ""
    xCampos(4, 0) = "idmon":          xCampos(4, 1) = TxtIdMon.Text:              xCampos(4, 2) = "S":    xCampos(4, 3) = "N":    xCampos(4, 4) = ""
    xCampos(5, 0) = "idunimed":       xCampos(5, 1) = TxtUnidad.Text:             xCampos(5, 2) = "S":    xCampos(5, 3) = "N":    xCampos(5, 4) = ""
    xCampos(6, 0) = "tippro":         xCampos(6, 1) = TxtTipPro.Text:             xCampos(6, 2) = "S":    xCampos(6, 3) = "N":    xCampos(6, 4) = ""
    xCampos(7, 0) = "idfam":          xCampos(7, 1) = TxtIdFamilia.Text:          xCampos(7, 2) = "S":    xCampos(7, 3) = "N":    xCampos(7, 4) = ""
    xCampos(8, 0) = "idclas":         xCampos(8, 1) = TxtIdClase.Text:            xCampos(8, 2) = "S":    xCampos(8, 3) = "N":    xCampos(8, 4) = ""
    xCampos(9, 0) = "idsubclas":      xCampos(9, 1) = TxtIdSubClase.Text:         xCampos(9, 2) = "S":    xCampos(9, 3) = "N":    xCampos(9, 4) = ""
    xCampos(10, 0) = "tipo":          xCampos(10, 1) = TxtIdTipmov.Text:          xCampos(10, 2) = "S":   xCampos(10, 3) = "N":   xCampos(10, 4) = ""
    
'    If EscribirNuevoRegistro(xCampos, "alm_inventario", xCon) = False Then
'        xCon.RollbackTrans
'        Exit Sub
'    End If
    
    Dim RstNew As New ADODB.Recordset
    
    RST_Busq RstNew, "SELECT * FROM alm_inventario", xCon
    RstNew.AddNew
    RstNew("id") = Str(xId)
    RstNew("codpro") = TxtCodPro.Text
    RstNew("descripcion") = TxtDescripcion.Text
    RstNew("desctec") = TxtDescripcion.Text
    RstNew("idmon") = TxtIdMon.Text
    RstNew("idunimed") = TxtUnidad.Text
    RstNew("tippro") = TxtTipPro.Text
    RstNew("idfam") = TxtIdFamilia.Text
    RstNew("idclas") = TxtIdClase.Text
    RstNew("idsubclas") = TxtIdSubClase.Text
    RstNew("tipo") = TxtIdTipmov.Text
    RstNew.Update
    
    xCon.CommitTrans
    MsgBox "El registro se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    xIdNewItem = xId
    Me.Hide
    Exit Sub

LaCague:
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el registro por el siguiente motivo: " & Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
End Sub

Private Sub CmdBusClase_Click()
    If VALIDAR_DATA(2) = False Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_clase.* FROM mae_clase WHERE mae_clase.idfam = " + CStr(Trim(TxtIdFamilia.Text)) + " ORDER BY mae_clase.descripcion ASC ;"
    
    xform.Titulo = "Buscando Clase"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        CODIGOTMP = NulosN(TxtIdClase.Text)
        
        TxtIdClase.Text = xRs("id")
        LblPrefijo2.Caption = xRs("prefijo")
        LblClase.Caption = xRs("descripcion")
        
        If CODIGOTMP <> 0 And CODIGOTMP <> NulosN(TxtTipPro.Text) Then
            TxtIdSubClase.Text = "":    LblSubClase.Caption = ""
        End If
        TxtIdSubClase.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Function VALIDAR_DATA(TIPO As Integer) As Boolean
Select Case TIPO
    Case 0 '--TIPO DE ITEM
    
    Case 1 '--FAMILIA
        If NulosN(Trim(TxtTipPro.Text)) = 0 Then
            MsgBox "Seleccione el Tipo de Item" + vbCr + "Luego Continue Seleccionando la Familia", vbExclamation, xTitulo
            'TabOne2.CurrTab = 0
            TxtTipPro.SetFocus
            Exit Function
        End If
    Case 2 '--CLASE
        If NulosN(Trim(TxtTipPro.Text)) = 0 Then
            MsgBox "Seleccione el Tipo de Item" + vbCr + "Luego Continue Seleccionando la Familia", vbExclamation, xTitulo
            'TabOne2.CurrTab = 0
            TxtTipPro.SetFocus
            Exit Function
        End If
        If NulosN(Trim(TxtIdFamilia.Text)) = 0 Then
            MsgBox "Seleccione la Familia" + vbCr + "Luego Continue Seleccionando la Clase", vbExclamation, xTitulo
            'TabOne2.CurrTab = 0
            TxtIdFamilia.SetFocus
            Exit Function
        End If
    Case 3 '--SUB CLASE
        If NulosN(Trim(TxtTipPro.Text)) = 0 Then
            MsgBox "Seleccione el Tipo de Item" + vbCr + "Luego Continue Seleccionando la Familia", vbExclamation, xTitulo
            'TabOne2.CurrTab = 0
            TxtTipPro.SetFocus
            Exit Function
        End If
        If NulosN(Trim(TxtIdFamilia.Text)) = 0 Then
            MsgBox "Seleccione la Familia" + vbCr + "Luego Continue Seleccionando la Clase", vbExclamation, xTitulo
            'TabOne2.CurrTab = 0
            TxtIdFamilia.SetFocus
            Exit Function
        End If
        If NulosN(Trim(TxtIdClase.Text)) = 0 Then
            MsgBox "Seleccione la Clase" + vbCr + "Luego Continue Seleccionando la Sub Clase", vbExclamation, xTitulo
            'TabOne2.CurrTab = 0
            TxtIdClase.SetFocus
            Exit Function
        End If
        
    Case 4 '--VALIDAR CUANDO SE GRABE O MODIFIQUE EL REGISTRO
        If NulosN(Trim(TxtTipPro.Text)) = 0 Then
            MsgBox "Seleccione el Tipo de Item" + vbCr + "Luego Continue Seleccionando la Familia", vbExclamation, xTitulo
            'TabOne2.CurrTab = 0
            TxtTipPro.SetFocus
            Exit Function
        End If
        If NulosN(Trim(TxtIdFamilia.Text)) = 0 Then
            MsgBox "Seleccione la Familia" + vbCr + "Luego Continue Seleccionando la Clase", vbExclamation, xTitulo
            'TabOne2.CurrTab = 0
            TxtIdFamilia.SetFocus
            Exit Function
        End If
        If NulosN(Trim(TxtIdClase.Text)) = 0 Then
            MsgBox "Seleccione la Clase" + vbCr + "Luego Continue Seleccionando la Sub Clase", vbExclamation, xTitulo
            'TabOne2.CurrTab = 0
            TxtIdClase.SetFocus
            Exit Function
        End If
        If NulosN(Trim(TxtIdSubClase.Text)) = 0 Then
            MsgBox "Seleccione la Sub Clase" + vbCr + "Luego Continue", vbExclamation, xTitulo
            'TabOne2.CurrTab = 0
            TxtIdSubClase.SetFocus
            Exit Function
        End If
    End Select
    VALIDAR_DATA = True
End Function


Private Sub CmdBusFam_Click()
    If VALIDAR_DATA(1) = False Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_familia.* FROM mae_familia where mae_familia.idtippro = " + CStr(Trim(TxtTipPro.Text)) + " ORDER BY mae_familia.descripcion ASC ; "
    
    xform.Titulo = "Buscando Familia"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        CODIGOTMP = NulosN(TxtIdFamilia.Text)
        
        TxtIdFamilia.Text = xRs("id")
        LblPrefijo1.Caption = xRs("prefijo") & ""
        LblFamilia.Caption = xRs("descripcion") & ""
        
        If CODIGOTMP <> 0 And CODIGOTMP <> NulosN(TxtIdFamilia.Text) Then
            LblFamilia.Caption = "":
            TxtIdClase.Text = "":       LblClase.Caption = ""
            TxtIdSubClase.Text = "":    LblSubClase.Caption = ""
        End If
        
        TxtIdClase.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMoneda_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_moneda.* FROM mae_moneda"
    
    xform.Titulo = "Buscando Moneda"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdMon.Text = xRs("id")
        LblMoneda.Caption = xRs("descripcion")
        TxtIdTipmov.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusSubClase_Click()
    If VALIDAR_DATA(3) = False Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_subclase.* FROM mae_subclase WHERE mae_subclase.idClas = " + Trim(TxtIdClase.Text) + " ORDER BY mae_subclase.descripcion ASC; "
    
    xform.Titulo = "Buscando Sub Clase"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdSubClase.Text = xRs("id")
        LblPrefijo3.Caption = xRs("prefijo") & ""
        LblSubClase.Caption = xRs("descripcion")
        TxtDescripcion.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipiTem_Click()
    'Dim xform As New EPS_Buscar.Buscar
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_tipoproducto.* FROM mae_tipoproducto"
    
    xform.Titulo = "Buscando Tipo de Producto"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        CODIGOTMP = NulosN(TxtTipPro.Text)
        TxtTipPro.Text = xRs("id")
        LblTipoPro.Caption = xRs("descripcion")
        LblPrefijo.Caption = xRs("prefijo")
        
        If CODIGOTMP <> 0 And CODIGOTMP <> NulosN(TxtTipPro.Text) Then
            TxtIdFamilia.Text = "":     LblFamilia.Caption = "":
            TxtIdClase.Text = "":       LblClase.Caption = ""
            TxtIdSubClase.Text = "":    LblSubClase.Caption = ""
        End If
        TxtIdFamilia.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipMovimiento_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_tipomovimiento.* FROM mae_tipomovimiento"
    
    xform.Titulo = "Buscando Tipo de Movimiento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdTipmov.Text = xRs("id")
        LblTipoMovi.Caption = xRs("descripcion")
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusUnidad_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Abreviatura":   xCampos(1, 1) = "abrev":          xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_unidades.* FROM mae_unidades"
    
    xform.Titulo = "Buscando Unidades"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        'LblIdUnidad.Caption = xRs("id")
        TxtUnidad.Text = xRs("id")
        LblDescUnidad.Caption = xRs("descripcion")
        TxtIdMon.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdCancel_Click()
    xIdProducto = 0
    Me.Hide
    xIdNewItem = 0
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        Blanquea
        TxtDescripcion.SetFocus
        TxtTipPro.SetFocus
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    QueHace = 1
    CaracteresNumericos = "0123456789." & Chr(8)
    xIdProducto = 0
End Sub

Sub Blanquea()
    TxtCodPro.Text = ""
    TxtTipPro.Text = ""
    TxtIdFamilia.Text = ""
    TxtIdClase.Text = ""
    TxtIdClase.Text = ""
    TxtDescripcion.Text = ""
    TxtUnidad.Text = ""
    TxtIdMon.Text = ""
    TxtIdTipmov.Text = ""
    TxtIdSubClase.Text = ""
    
    LblTipoPro.Caption = ""
    LblFamilia.Caption = ""
    LblClase.Caption = ""
    LblSubClase.Caption = ""
    LblDescUnidad.Caption = ""
    LblMoneda.Caption = ""
    LblTipoMovi.Caption = ""
End Sub

Private Sub TxtIdClase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtIdClase.Text <> "" Then
            If VALIDAR_DATA(2) = False Then Exit Sub
            Set RstTem = BuscaConCriterio("SELECT * FROM mae_clase WHERE id =" & NulosN(TxtIdClase.Text) & " AND mae_clase.idfam=" + CStr(Trim(TxtIdFamilia.Text)), xCon)
            If RstTem.RecordCount <> 0 Then
                LblClase.Caption = RstTem("descripcion") & ""
                LblPrefijo2.Caption = RstTem("prefijo") & ""
                TxtIdSubClase.Text = "":    LblSubClase.Caption = ""
            Else
                LblClase.Caption = ""
                TxtIdClase.Text = "":   LblClase.Caption = ""
            End If
        End If
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdClase_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusClase_Click
    End If
End Sub

Private Sub TxtIdFamilia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtIdFamilia.Text <> "" Then
            If VALIDAR_DATA(1) = False Then Exit Sub
            Set RstTem = BuscaConCriterio("SELECT * FROM mae_familia WHERE id =" & NulosN(TxtIdFamilia.Text) & " AND mae_familia.idtippro = " + CStr(Trim(TxtTipPro.Text)), xCon)
            If RstTem.RecordCount <> 0 Then
                LblFamilia.Caption = RstTem("descripcion") & ""
                LblPrefijo1.Caption = NulosC(RstTem("prefijo")) & ""
            Else
                LblFamilia.Caption = ""
                TxtIdFamilia.Text = ""
            End If
        End If
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdFamilia_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusFam_Click
    End If
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtIdMon.Text <> "" Then
            Set RstTem = BuscaConCriterio("SELECT * FROM mae_moneda WHERE id =" & NulosN(TxtIdMon.Text) & "", xCon)
            If RstTem.RecordCount <> 0 Then
                LblMoneda.Caption = RstTem("descripcion")
            Else
                TxtIdMon.Text = ""
                LblMoneda.Caption = ""
            End If
        End If
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMoneda_Click
    End If
End Sub

Private Sub TxtIdSubClase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtIdSubClase.Text <> "" Then
            If VALIDAR_DATA(3) = False Then Exit Sub
            Set RstTem = BuscaConCriterio("SELECT * FROM mae_subclase WHERE id =" & NulosN(TxtIdSubClase.Text) & " AND mae_subclase.idClas = " + Trim(TxtIdClase.Text), xCon)
            If RstTem.RecordCount <> 0 Then
                LblSubClase.Caption = RstTem("descripcion")
                LblPrefijo3.Caption = RstTem("prefijo")
            Else
                LblSubClase.Caption = ""
                TxtIdSubClase.Text = ""
            End If
        End If
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdSubClase_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusSubClase_Click
    End If
End Sub

Private Sub TxtIdTipmov_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtIdTipmov.Text <> "" Then
            Set RstTem = BuscaConCriterio("SELECT * FROM mae_tipomovimiento WHERE id =" & NulosN(TxtIdTipmov.Text) & "", xCon)
            If RstTem.RecordCount <> 0 Then
                LblTipoMovi.Caption = RstTem("descripcion") & ""
                TxtIdTipmov.Text = RstTem("id") & ""
            Else
                LblTipoMovi.Caption = ""
                TxtIdTipmov.Text = ""
            End If
        End If
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdTipmov_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipMovimiento_Click
    End If
End Sub

Private Sub TxtTipPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtTipPro.Text <> "" Then
            Set RstTem = BuscaConCriterio("SELECT * FROM mae_tipoproducto WHERE id =" & NulosN(TxtTipPro.Text) & "", xCon)
            If RstTem.RecordCount <> 0 Then
                LblTipoPro.Caption = RstTem("descripcion") & ""
                LblPrefijo.Caption = RstTem("prefijo") & ""
                TxtIdFamilia.Text = "":     LblFamilia.Caption = "":
                TxtIdClase.Text = "":       LblClase.Caption = ""
                TxtIdSubClase.Text = "":    LblSubClase.Caption = ""
            Else
                LblTipoPro.Caption = ""
                TxtTipPro.Text = ""
            End If
        End If
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipPro_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipiTem_Click
    End If
End Sub

Private Sub TxtUnidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TxtUnidad.Text <> "" Then
            Set RstTem = BuscaConCriterio("SELECT * FROM mae_unidades WHERE id =" & NulosN(TxtUnidad.Text) & "", xCon)
            If RstTem.RecordCount <> 0 Then
                LblDescUnidad.Caption = RstTem("descripcion") & ""
            Else
                LblDescUnidad.Caption = ""
                TxtUnidad.Text = ""
            End If
        End If
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtUnidad_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusUnidad_Click
    End If
End Sub

