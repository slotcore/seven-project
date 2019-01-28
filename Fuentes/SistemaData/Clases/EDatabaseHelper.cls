VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EDatabaseHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mProvider As String
Private mPassword As String
Private mUserId As String
Private mDataSource As String
Private mGrupoTrabajo As String

Public Property Get Provider() As String
    Provider = mProvider
End Property
Public Property Let Provider(ByVal NewValue As String)
    mProvider = NewValue
End Property

Public Property Get Password() As String
    Password = mPassword
End Property
Public Property Let Password(ByVal NewValue As String)
    mPassword = NewValue
End Property

Public Property Get UserId() As String
    UserId = mUserId
End Property
Public Property Let UserId(ByVal NewValue As String)
    mUserId = NewValue
End Property

Public Property Get DataSource() As String
    DataSource = mDataSource
End Property
Public Property Let DataSource(ByVal NewValue As String)
    mDataSource = NewValue
End Property

Public Property Get GrupoTrabajo() As String
    GrupoTrabajo = mGrupoTrabajo
End Property
Public Property Let GrupoTrabajo(ByVal NewValue As String)
    mGrupoTrabajo = Trim(NewValue)
End Property

Function GetDbProvider() As String
    GetDbProvider = mProvider
End Function

'Function GetDbConnectionString() As String
'    Dim xCadConexion As String
'    xCadConexion = "Provider = " + mProvider _
'                    & ";Password = " & mPassword & "; Persist Security Info = true" _
'                    & ";User ID = " & mUserId & "; Data Source = " + DataSource _
'                    & "; Jet OLEDB:System database = " + mGrupoTrabajo
'
'    GetDbConnectionString = xCadConexion
'End Function

Function GetDbConnectionString(RutaDataSource As String) As String
    Dim xCadConexion As String
    Me.DataSource = RutaDataSource
    xCadConexion = "Provider = " + Me.Provider _
                    & ";Password = " & Me.Password & "; Persist Security Info = true" _
                    & ";User ID = " & Me.UserId & "; Data Source = " + Me.DataSource _
                    & "; Jet OLEDB:System database = " + Me.GrupoTrabajo
    
    GetDbConnectionString = xCadConexion
End Function

Private Sub Class_Initialize()
    Dim AP_RUTASY As String
    Dim AP_RUTABD As String
    Dim AP_USER As String
    Dim AP_PASS As String
    Dim F As New SistemaData.Funciones
        
    AP_RUTASY = F.LeerLineaINI(Trim(App.Path) + "\app.ini", "RUTASY", "RUTAS")
'    AP_RUTABD = F.LeerLineaINI(Trim(App.Path) + "\app.ini", "RUTABD", "RUTAS")
    AP_USER = F.LeerLineaINI(Trim(App.Path) + "\app.ini", "USER", "CREDENCIALES")
    AP_PASS = F.LeerLineaINI(Trim(App.Path) + "\app.ini", "PASS", "CREDENCIALES")
    
'    Me.DataSource = AP_RUTABD + "data.mdb"
    Me.GrupoTrabajo = AP_RUTASY + "app.mdw"
    Me.UserId = F.Desencriptar(AP_USER)
    Me.Password = F.Desencriptar(AP_PASS)
    Me.Provider = "Microsoft.Jet.OLEDB.4.0"
End Sub

