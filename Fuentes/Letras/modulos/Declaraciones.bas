Attribute VB_Name = "Declaraciones"
Option Explicit

Public xCon As New ADODB.Connection
Public xTitulo As String
Public xNomEmp, xNumRuc As String
Public CaracteresNumericos As String

Public NomEmp As String
Public NumRUC As String
Public CONTABILIZAR As Boolean
Public xMes As Integer
Public AnoTra As String
Public xIdEmpresa As Integer

Public xDeDonde As Integer      ' almacenara el un valor para saber de donde se esta invocando a la librerias si del menu de compras o del menu de opciones
                                ' 1 = menu compras
                                ' 2 = menu de opciones

Public xIdUsuario As Integer


Global xIGV As Double
Global xRetencion As Double
    
Public xIdMenu As Integer                ' Especifica el id del menu(formulario que accede el usuario)


Function AbrirOtraConeccion(Ruta As String) As ADODB.Connection
    'Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xFun As New eps_librerias.FuncionesData
    
    Dim Con As New ADODB.Connection
    Dim xCad As String
        
    xCad = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & Trim(Ruta) & "'; " _
    & " Persist Security Info=False;Jet OLEDB:Database"
    
    Con.ConnectionString = xCad
    If Con.State = 0 Then
        Con.Open
        Set AbrirOtraConeccion = Con
    End If
End Function

Sub CargaDatosEmpresa()
    Dim Rst As New ADODB.Recordset
    
    RST_Busq Rst, "SELECT * FROM mae_empresa", xCon
    NomEmp = NulosC(Rst("nomemp"))
    NumRUC = NulosC(Rst("numruc"))
    CONTABILIZAR = NulosN(Rst("procon"))
    AnoTra = NulosC(Rst("anotra"))
    
    Set Rst = Nothing
    
    RST_Busq Rst, "select tasa from mae_impuestos where id=1", xCon
    If Rst.State = 1 Then
        If Rst.RecordCount <> 0 Then
            xIGV = NulosN(Rst("tasa")) / 100
        Else
            xIGV = 0.19
        End If
    Else
        xIGV = 0.19
    End If
    Set Rst = Nothing
    '--definiendo los porcentajes como constantes
    
    xRetencion = 0.06

    
End Sub




