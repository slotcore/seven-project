VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar PG 
      Height          =   330
      Left            =   540
      TabIndex        =   1
      Top             =   1785
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   300
      Left            =   1470
      TabIndex        =   0
      Top             =   810
      Width           =   3030
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    Dim A, B As Integer
    Dim RstCompraSEVEN As New ADODB.Recordset
    Dim RstDiario As New ADODB.Recordset
    Dim RstBuscaDiario As New ADODB.Recordset
    Dim RstBuscaRegCom As New ADODB.Recordset
    
    RST_Busq RstCompraSEVEN, "SELECT com_compras.*, con_tc.impven, * FROM com_compras LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
        & " WHERE (((com_compras.numreg) Like '07%'))", xCon

    RST_Busq RstDiario, "SELECT * FROM con_diario", xCon
    
    
    If RstCompraSEVEN.RecordCount <> 0 Then
        
        PG.Max = RstCompraSEVEN.RecordCount
        
        RstCompraSEVEN.MoveFirst
        
        For A = 1 To RstCompraSEVEN.RecordCount
            PG.Value = A
            
            xCon.Execute "DELETE * FROM con_diario WHERE idmov = " & RstCompraSEVEN("com_compras.id") & " AND idlib = 1 and numasi = '" & Mid(RstCompraSEVEN("numreg"), 3, 4) & "'"
            Set RstBuscaRegCom = Nothing
            Set RstBuscaRegCom = Nothing
            RST_Busq RstBuscaRegCom, "SELECT CON_Registro_Compras.NroCompra, CON_Registro_Compras.numser, CON_Registro_Compras.numdoc, CON_Registro_Compras.Libro" _
                & " From CON_Registro_Compras WHERE (((CON_Registro_Compras.NroCompra) Like '07%') AND ((CON_Registro_Compras.numser)='" & RstCompraSEVEN("numser") & "') " _
                & " AND ((CON_Registro_Compras.numdoc)='" & RstCompraSEVEN("numdoc") & "'))", xCon
            
            If RstBuscaRegCom.RecordCount <> 0 Then
                RST_Busq RstBuscaDiario, "SELECT CON_Diario_General.Nro_Libro, CON_Diario_General.Cuenta, Sum(CON_Diario_General.Monto_Debe) AS SumaDeMonto_Debe, " _
                    & " Sum(CON_Diario_General.Monto_Haber) AS SumaDeMonto_Haber, CON_Diario_General.Cod_Moneda, CON_Diario_General.Libro   " _
                    & " From CON_Diario_General " _
                    & " GROUP BY CON_Diario_General.Nro_Libro, CON_Diario_General.Cuenta, CON_Diario_General.Cod_Moneda, CON_Diario_General.Libro " _
                    & " HAVING (((CON_Diario_General.Nro_Libro)='" & RstBuscaRegCom("NroCompra") & "') AND ((CON_Diario_General.Libro)='" & RstBuscaRegCom("Libro") & "'))", xCon

                If RstBuscaDiario.RecordCount <> 0 Then
                    RstBuscaDiario.MoveFirst
                    
                    For B = 1 To RstBuscaDiario.RecordCount
                        RstDiario.AddNew
                        RstDiario("año") = 2007
                        RstDiario("idmes") = 7     'cambiar este campo para los demas meses
                        RstDiario("idlib") = 1
                        RstDiario("idmov") = RstCompraSEVEN("com_compras.id")
                        RstDiario("idcue") = Busca_Codigo(RstBuscaDiario("Cuenta"), "cuenta", "id", "con_planctas", "C", xCon)
                        RstDiario("numasi") = Mid(RstCompraSEVEN("numreg"), 3, 4)
                        RstDiario("tc") = RstCompraSEVEN("impven")
                        RstDiario("fchasi") = "01/07/07"  'cambiar este campo para los demas meses
                        If RstCompraSEVEN("com_compras.idmon") = 1 Then
                            RstDiario("imphabsol") = RstBuscaDiario("sumademonto_haber")
                            RstDiario("impdebsol") = RstBuscaDiario("sumademonto_debe")
                            
                            RstDiario("imphabdol") = 0
                            RstDiario("impdebdol") = 0
                        Else
                            RstDiario("imphabsol") = RstBuscaDiario("sumademonto_haber") * RstCompraSEVEN("impven")
                            RstDiario("impdebsol") = RstBuscaDiario("sumademonto_debe") * RstCompraSEVEN("impven")
    
                            RstDiario("imphabdol") = RstBuscaDiario("sumademonto_haber")
                            RstDiario("impdebdol") = RstBuscaDiario("sumademonto_debe")
                        End If
                        
                        RstDiario.Update
                        
                        RstBuscaDiario.MoveNext
                        If RstBuscaDiario.EOF = True Then
                            Exit For
                        End If
                    Next B
                End If
            End If
            RstCompraSEVEN.MoveNext
            If RstCompraSEVEN.EOF = True Then Exit For
        Next A
    End If
    MsgBox "Termine rendida"
End Sub
