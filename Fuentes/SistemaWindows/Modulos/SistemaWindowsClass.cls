VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SistemaWindowsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Private mLoadingForm As New LoadingForm
Private mProgressForm As New ProgressForm

Public Enum TipoMensaje
    Error = 1
    Warning = 2
End Enum

Sub ShowProgress(Proceso As String, Min As Long, Max As Long)
    mProgressForm.Proceso = Proceso
    mProgressForm.Min = Min
    mProgressForm.Max = Max
    mProgressForm.Show
End Sub

Sub SetProgress(DescriptionProgress As String, Value As Long)
    mProgressForm.Descripcion = DescriptionProgress
    mProgressForm.Value = Value
    mProgressForm.Refresh
End Sub

Sub SetSubProgress(SubProceso As String, SubDescriptionProgress As String)
    mProgressForm.SubProceso = SubProceso
    mProgressForm.SubDescripcion = SubDescriptionProgress
    mProgressForm.Refresh
End Sub

Sub HideProgress()
    Unload mProgressForm
End Sub

Sub ShowLoading(view As Boolean)
'    DoEvents
'    mLoadingForm.ZOrder 0
'    If view Then mLoadingForm.Show Else Unload mLoadingForm
'    mLoadingForm.Refresh
End Sub

Sub ShowMessage(tipo As TipoMensaje)
    Dim M As New MensajeBoxForm
    M.TipoMensaje = tipo
    M.Show
End Sub
