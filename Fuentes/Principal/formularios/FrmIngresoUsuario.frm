VERSION 5.00
Begin VB.Form FrmIngresoUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   Icon            =   "FrmIngresoUsuario.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   3825
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   1935
      TabIndex        =   3
      Top             =   1410
      Width           =   1095
   End
   Begin VB.CommandButton CmdAcepta 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   795
      TabIndex        =   2
      Top             =   1410
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   1365
      Left            =   30
      TabIndex        =   4
      Top             =   -45
      Width           =   3765
      Begin VB.TextBox TxtPass 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1635
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox TxtUsuario 
         Height          =   300
         Left            =   1635
         TabIndex        =   0
         Top             =   375
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Contraseña"
         Height          =   225
         Left            =   435
         TabIndex        =   6
         Top             =   750
         Width           =   930
      End
      Begin VB.Label Label2 
         Caption         =   "Usuario"
         Height          =   225
         Left            =   435
         TabIndex        =   5
         Top             =   420
         Width           =   930
      End
   End
End
Attribute VB_Name = "FrmIngresoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo    : FRMMANOPCIONESUSUARIO
'* Tipo              : FORMULARIO
'* Descripcion       : CONTROLA EL INGRESO DE LOS USUARIOS AL SISTEMA, VALIDA EL USUARIO Y EL PASSWORD
'*                     , CARGANDO LAS OPCIONES DE ACCESO AL MENU DEL SISTEMA
'* DISEÑADO POR      : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION   : 04/09/09
'* VERSION           : 1.0
'*****************************************************************************************************
Option Explicit
Dim xContador As Integer

Private Sub CmdAcepta_Click()
    ' VERIFICAMOS QUE EL CONTADOR DE OPORTUNIDADES DE INGRESO NO HAYA EXCEDIDO EL LIMITE, DE HABER
    ' EXCEDIDO EL SISTEMA NO PERMITIRA LA EJECUCION DEL SISTEMA EVITANDO ASI EL INGRESO DE USUARIOS NO
    ' AUTORIZADOS.
    ' EN ESTE PROCEDIMIENTO TAMBIEN SE VALIDA QUE NO SE INGRESEN VALORES NULOS PARA EL USUARIO Y EL
    ' PASSWORD
    
    If xContador = 3 Then
        MsgBox "Ha excedido el limite de oportunidades para el ingreso del usuario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Unload Me
        xCon.Close
        Set xCon = Nothing
        End
    End If
    
    If TxtUsuario.Text = "" Then
        MsgBox "No ha especificado el usuario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtUsuario.SetFocus
        Exit Sub
    End If
    
    If TxtPass.Text = "" Then
        MsgBox "No ha especificado el password", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtPass.SetFocus
        Exit Sub
    End If
        
    Dim RstUsu As New ADODB.Recordset
    
    ' BUSCAMOS EL USUARIO INGRESADO, CONSULTANDO EL USUARIO Y PASSWORD INGRESADO
    RST_Busq RstUsu, "SELECT * FROM mae_usuarios WHERE login = '" & Trim(TxtUsuario.Text) & "'" _
        & " AND pass = '" & Trim(TxtPass.Text) & "'", xCon
    
    If RstUsu.RecordCount = 0 Then
        ' SI EL USUARIO NO ES ENCONTRADO INCREMENTAMOS EL CONTADOR DE OPORTUNIDADES DE INGRES EN UNO
        xContador = xContador + 1
        MsgBox "Usuario o contraseña incorrectos, intente de nuevo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtUsuario.Text = ""
        TxtPass.Text = ""
        TxtUsuario.SetFocus
    Else
        ' SI EL USUARIO ES ENCONTRADO, SE VERIFICA QUE SEA UN USUARIO ACTIVO
        If RstUsu("activo") = 0 Then
            ' DE NO SER UN USUARIO ACTIVO AVISAMOS Y PREPARMOS PARA EL INGRESO DE OTRO USUARIO Y SU RESPECTIVO PASSWORD
            MsgBox "El usuario " + Trim(TxtUsuario.Text) + " ha sido desactivado ingrese otro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtUsuario.Text = ""
            TxtPass.Text = ""
            TxtUsuario.SetFocus
            Exit Sub
        End If
        
        ' CONFIRMAMOS EL INGRESO DEL USUARIO AL SISTEMA
        xIdUsuario = Val(RstUsu("id"))
        MsgBox "Se registro el ingreso del usuario: " & RstUsu("nomusu"), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        MDIPrincipal.StatusBar1.Panels(2).Text = "Usuario  : " + Trim(RstUsu("login"))
        
'        ' CARGAMOS LAS OPCIONES DISPONIBLES DEL MENU PARA EL USUARIO
'        SetearMenus xIdUsuario
        ' ACTIVAMOS LAS OPCIONES DEL MENU10 (MENU DE CONFIGURACION DEL SISTEMA)
        
''        ActivarOpcionesMenu10
        
        PedirUsuario = False ' Indicamos al sistema que yan o pida usuario
        Unload Me
    End If
End Sub

''*****************************************************************************************************
''* Nombre Modulo  : SetearMenus()
''* Tipo           : PROCEDIMIENTO
''* Descripcion    : CARGA LAS OPCIONES DEL MENU DISPONIBLES PARA EL USUARIO ACTUAL
''* Paranetros     : NOMBRE    |  TIPO     |  DESCRIPCION
''*                  ----------------------------------------------------------------------------------
''*                  Idusuario | INTEGER   | CODIGO UNICO DEL USUARIO ACTUAL
''* Retorna        : NULL
''*****************************************************************************************************
'Sub SetearMenus(Idusuario As Integer)
'    Dim Rst As New ADODB.Recordset
'    Dim A As Integer
'    On Error Resume Next
'
'    ' CARGAMOS LAS OPCIONES DEL MENU DISPONIBLES PARA EL USUARIO
'    RST_Busq Rst, "SELECT mae_menu.id, mae_menu.tipo, mae_menu.descripcion, mae_menu.nomcon, mae_menuusuario.opcion1, " _
'        & " mae_menuusuario.opcion2, mae_menuusuario.opcion3, mae_menuusuario.acceso FROM mae_menu LEFT JOIN mae_menuusuario ON " _
'        & " mae_menu.id = mae_menuusuario.idmenu WHERE (((mae_menuusuario.idusuario)= " & Idusuario & "))", xCon
'
'    ' PREGUNTAMOS SI HAY OPCIONES DISPONIBLES
'    If Rst.RecordCount <> 0 Then
'        Rst.MoveFirst
'        'RECORREMOS LAS OPCIONES DISPONIBLES PARA SU RESPECTIVA ACTIVACION
'        For A = 1 To Rst.RecordCount
'            MDIPrincipal(Rst("nomcon")).Enabled = Rst("acceso")
'            Rst.MoveNext
'            If Rst.EOF = True Then Exit For
'        Next A
'    End If
'    Err.Clear
'End Sub

Private Sub CmdCancel_Click()
    MsgBox "No has ingresado ningun usuario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set xCon = Nothing
    Unload Me
    End
End Sub

Private Sub Form_Load()
    TxtUsuario.Text = ""
    TxtPass.Text = ""
    xContador = 0
End Sub

Private Sub TxtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        
        If TxtUsuario.Text <> "" And TxtPass.Text <> "" Then
            CmdAcepta_Click
        End If
    End If
End Sub

Private Sub TxtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub
