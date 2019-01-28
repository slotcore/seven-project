Attribute VB_Name = "Principal"
Public xCon As New ADODB.Connection
Public xTitulo As String

Sub Main()
    Dim xFun As New eps_librerias.FuncionesData
    xFun.F_BASEDATOS = "C:\seven\data\2008\0001\data.mdb"
    xFun.F_GRUPOTRABAJO = "c:\seven\seven.mdw"
    xFun.F_PASSWORD = "010419762005"
    xFun.F_USUARIO = "cav2005sialp"
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"
    Set xCon = xFun.AbrirConeccion
    If xCon.State = 1 Then
        Form1.Show
    Else
        MsgBox "No se pudo abri la conecion a la BD"
    End If
End Sub
