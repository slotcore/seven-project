VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EDataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mDataSource As String
Private mConexion As New ADODB.Connection
Private mCommand As New ADODB.Command
Private rstRecordSet As New ADODB.Recordset
Private mCommandText As String
Private mProcedureName As String

Public Property Get DataSource() As String
    DataSource = mDataSource
End Property
Public Property Let DataSource(ByVal NewValue As String)
    mDataSource = NewValue
End Property

Public Property Let CommandText(ByVal NewValue As String)
    mCommand.ActiveConnection = mConexion
    mCommand.CommandType = adCmdText
    mCommand.CommandText = NewValue
End Property

Public Property Let ProcedureName(ByVal NewValue As String)
    mCommand.ActiveConnection = mConexion
    mCommand.CommandType = adCmdStoredProc
    mCommand.CommandText = NewValue
End Property

Public Property Get Connection() As ADODB.Connection
    Set Connection = mConexion
End Property
Public Property Set Connection(ByVal NewValue As ADODB.Connection)
    Set mConexion = NewValue
End Property

Public Sub AddParameter(parameterName As String, parameterType As DataTypeEnum, parameterDirection As ParameterDirectionEnum, parameterValue As Variant, Optional sizeValue As Long = 0)
    Dim P As New ADODB.Parameter
    
    If (parameterType = adVarChar) Then
        mCommand.Parameters.Append mCommand.CreateParameter(parameterName, parameterType, parameterDirection, 255)
    ElseIf (parameterType = adChar) Then
        mCommand.Parameters.Append mCommand.CreateParameter(parameterName, parameterType, , sizeValue)
    Else
        mCommand.Parameters.Append mCommand.CreateParameter(parameterName, parameterType)
    End If
    'mCommand.Parameters.Append P
    mCommand(parameterName) = parameterValue
    
'    P.Name = parameterName
'    P.Type = parameterType
'    P.Direction = parameterDirection
'    P.Value = parameterValue
'    P.Size = sizeValue
'    mCommand.Parameters.Append P
End Sub

Public Sub ClearParameter()
    Dim lngX As Long
    For lngX = (mCommand.Parameters.Count - 1) To 0 Step -1
        mCommand.Parameters.Delete lngX
    Next
End Sub

Public Function GetRecordset(Optional ForwardOnly As Boolean = True) As ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    
    With xRs
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .Open mCommand
    End With
    
    'Set GetRecordset = mCommand.Execute
    Set GetRecordset = xRs
End Function

Public Function GetConnection() As ADODB.Connection
    Set GetConnection = mConexion
End Function

Public Sub Execute()
    mCommand.Execute
End Sub

Public Sub BeginTrans()
    If mConexion.State = 0 Then mConexion.Open
    mConexion.BeginTrans
End Sub

Public Sub CommitTrans()
    mConexion.CommitTrans
End Sub

Public Sub RollbackTrans()
    mConexion.RollbackTrans
End Sub

'Private Sub OpenConnection()
'    Dim C As New EDatabaseHelper
'
'    mConexion.ConnectionString = C.GetDbConnectionString
'    mConexion.Open
'End Sub

Public Sub OpenConnection(DataSource As String)
    Dim C As New EDatabaseHelper
    mDataSource = DataSource
    mConexion.ConnectionString = C.GetDbConnectionString(DataSource)
    mConexion.Open
End Sub

Public Sub CloseConnection()
    If (mConexion.State = adStateOpen) Then
        mConexion.Close
    End If
End Sub

'Private Sub Class_Initialize()
'    OpenConnection
'End Sub

Private Sub Class_Terminate()
    If mDataSource <> "" Then
        CloseConnection
    End If
End Sub

