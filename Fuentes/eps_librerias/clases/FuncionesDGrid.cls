VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FuncionesDGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public xNumRuc As String
Public xNomEmp As String

Sub ExportarDGExcel(xRst As ADODB.Recordset, xDataGrid As Object, xTitulo As String)
    Dim A, B, xFila, xCol, xColExcel As Integer
    On Error GoTo error
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    
    objExcel.Visible = True
    'determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1
    
    'abre el Libro
    objExcel.WindowState = 1
    objExcel.Workbooks.Add
    
    'Set xRst.ActiveConnection = Nothing
    'Set xDataGrid.DataSource = xRst
        
    With objExcel.ActiveSheet
        xRst.MoveFirst
        
        .Cells(1, 1) = "EMPRESA : " & xNomEmp
        .Cells(2, 1) = "Nº R.U.C. : " & xNumRuc
        
        .Cells(4, 1) = xTitulo
        xFila = 6
        xCol = 1
        
        For A = 1 To xRst.RecordCount
            xCol = 1
            ' IMPRIMIMOS LOS TITULOS DEL GRID
            If A = 1 Then
                xColExcel = 1
                For B = 0 To xDataGrid.Columns.Count - 1
                    .Cells(xFila, xCol) = "'" & F_NulosC(xDataGrid.Columns(B).Caption)
                    .Columns(xColExcel).ColumnWidth = (xDataGrid.Columns(B).Width / 100)
                    xCol = xCol + 1
                    xColExcel = xColExcel + 1
                Next B
                xFila = xFila + 1
                xCol = 1
            End If
            
            For B = 0 To xDataGrid.Columns.Count - 1
                'If xRst(xDataGrid.Columns(B).DataField).Type = 202 Then  ' si es texto
                If xRst.Fields(xDataGrid.Columns(B).DataField).Type = 202 Or xRst.Fields(xDataGrid.Columns(B).DataField).Type = 203 Then   ' SI ES TEXTO
                    If F_NulosC(xRst(xDataGrid.Columns(B).DataField)) = "" Then
                        .Cells(xFila, xCol) = ""
                    Else
                        .Cells(xFila, xCol) = "'" & F_NulosC(xRst(xDataGrid.Columns(B).DataField))
                    End If
                End If
                
                If xRst.Fields(xDataGrid.Columns(B).DataField).Type = 7 Then     ' SI ES UNA FECHA
                    If F_NulosC(xRst(xDataGrid.Columns(B).DataField)) = "" Then
                        .Cells(xFila, xCol) = ""
                    Else
                        .Cells(xFila, xCol) = F_NulosC(xRst(xDataGrid.Columns(B).DataField))
                    End If
                End If
                
                If xRst.Fields(xDataGrid.Columns(B).DataField).Type = 5 Then     ' SI ES UNA NUMERO
                    If F_NulosC(xRst(xDataGrid.Columns(B).DataField)) = "" Then
                        .Cells(xFila, xCol) = ""
                    Else
                        .Cells(xFila, xCol) = F_NulosC(xRst(xDataGrid.Columns(B).DataField))
                    End If
                End If
                
                xCol = xCol + 1
            Next B
            xRst.MoveNext
            If xRst.EOF = True Then
                Exit For
            End If
            xFila = xFila + 1
        Next A
    End With
    
    MsgBox "El proceso de exportación terminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
    
error:
    'Resume
    Set objExcel = Nothing
End Sub
