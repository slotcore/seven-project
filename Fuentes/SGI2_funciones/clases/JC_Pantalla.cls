VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JC_Pantalla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'---------1*********2---------3*********4---------5*********6---------7*********8----5----9
'Resolucion.BAS, Modulo para ajustar los controles en un formulario a travez de
'distintos tamaños de pantalla y al cambiar el tamaño del mismo.
'Fecha de Creacion:14/04/2004
'Agregar el siguiente codigo al formulario:
'
'Private Sub Form_Load()
'  Ajusta Me
'End Sub
'
'If Me.WindowState <> vbMinimized Then
'   ReAjusta Me
'
'End If
'Modulo adaptado de:
'182070 - HOWTO Create a Resolution-Independent Form
'http://support.microsoft.com/default.aspx?scid=kb;en-us;182070
'187562 - HOWTO Resize the Controls in SSTab When Form is Resized
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;187562
'
'Adaptado por: Hector Garcia ®, heichito@hotmail.com
'
'ADVERTENCIA: Este Modulo representa una implementacion publica para _
'multiples aplicaciones. Cualquier modificacion a Funciones, Tipos de datos _
'o variables podria provocar una falla en los otros sistemas


Option Explicit

Type ctrObj
    Name As String
    Index As Long
    Parrent As String
    Top As Long
    Left As Long
    Height As Long
    Width As Long
    ScaleHeight As Long
    ScaleWidth As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type FRMSIZE
    Height As Long
    Width As Long
End Type

Private Const SPI_GETWORKAREA = 48
'Public MyFormSize As FRMSIZE
Dim MyFormSize As FRMSIZE
Public DesignX As Long
Public DesignY As Long

Public Xtwips As Integer, Ytwips As Integer
Public Xpixels As Integer, Ypixels As Integer


Public RePosForm As Boolean
Public DoResize As Boolean


'Obtener o establecer diversas configuraciones de sistema, tamaño de pantalla activa
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" _
        (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, _
        ByVal fuWinIni As Long) As Long


Private FormRecord() As ctrObj
Private ControlRecord() As ctrObj
Private MaxForm As Long
Private MaxControl As Long





Public Function ObtenAreaTrabajo() As RECT
    Dim rct As RECT
    Dim lRet As Long
    lRet = SystemParametersInfo(SPI_GETWORKAREA, vbNull, rct, 0)
    If lRet > 0 Then ObtenAreaTrabajo = rct
End Function

Private Sub Resize_For_Resolution(ByVal SFX As Single, _
                          ByVal SFY As Single, MyForm As Form)
    Dim i As Long
    Dim SFFont As Single
    Dim Ctl As Control
    Dim CtlCln As New Collection
    
    SFFont = (SFX + SFY) / 2  ' average scale
    ' Size the Controls for the new resolution
    On Error Resume Next  ' for read-only or nonexistent properties
    
    With MyForm
        For i = 0 To .Count - 1
            If TypeName(.Controls(i)) <> "CommonDialog" And _
             TypeName(.Controls(i)) <> "Menu" And _
             TypeName(.Controls(i)) <> "Timer" And _
             TypeName(.Controls(i)) <> "Winsock" And _
             TypeName(.Controls(i)) <> "ImageList" Then
                'El control se encuentra en una ficha oculta de  ssTab?
                'Adaptacion de MSDN KB ID : 187562
                If .Controls(i).Left < 0 Then
                    CtlCln.Add .Controls(i)
                    .Controls(i).Left = .Controls(i).Left + 75000
                End If
            End If
'            "VSFlexGrid",
            Select Case TypeName(.Controls(i))
                Case "ComboBox", "TextBox", "DTPicker", "CheckBox", "OptionButton", "TextBoxFecha"
                    .Controls(i).Left = .Controls(i).Left * SFX
                    .Controls(i).Top = .Controls(i).Top * SFY
                    .Controls(i).Width = .Controls(i).Width * SFX
                Case "CommandButton", "Label"
                    .Controls(i).Left = .Controls(i).Left * SFX
                    .Controls(i).Top = .Controls(i).Top * SFY
                    .Controls(i).Width = .Controls(i).Width * SFX
                Case "Line"
                    .Controls(i).X1 = .Controls(i).X1 * SFX
                    .Controls(i).X2 = .Controls(i).X2 * SFX
                    .Controls(i).Y1 = .Controls(i).Y1 * SFY
                    .Controls(i).Y2 = .Controls(i).Y1
                Case "Splitter"
                               
                Case Else
                .Controls(i).Move .Controls(i).Left * SFX, _
                .Controls(i).Top * SFY, _
                .Controls(i).Width * SFX, _
                .Controls(i).Height * SFY
    
            End Select
            
            ' Be sure to resize and reposition before changing the FontSize
            '.Controls(I).FontSize = .Controls(I).FontSize * SFFont
        Next i
        'Regresa la posicion original del control dentro del ssTab
        'Adaptacion de MSDN KB ID : 187562
        For Each Ctl In CtlCln
            Ctl.Left = Ctl.Left - 75000
        Next Ctl
        If RePosForm Then
            ' Now size the Form
            .Move .Left * SFX, .Top * SFY, .Width * SFX, .Height * SFY
        End If
    End With
End Sub


Public Sub Ajusta(ByVal MyForm As Object)
    Dim ScaleFactorX As Single, ScaleFactorY As Single  ' Scaling factors
    Dim ScaleMode As Single

    ' Size of Form in Pixels at design resolution
    DesignX = 800
    DesignY = 600
    RePosForm = True   ' Flag for positioning Form
    DoResize = False   ' Flag for Resize Event
    ' Set up the screen values
    Xtwips = Screen.TwipsPerPixelX
    Ytwips = Screen.TwipsPerPixelY
    Ypixels = Screen.Height / Ytwips ' Y Pixel Resolution
    Xpixels = Screen.Width / Xtwips  ' X Pixel Resolution
    
    If DesignX = Xpixels Then Exit Sub
    
    
    ' Determine scaling factors
    ScaleFactorX = (Xpixels / DesignX)
    ScaleFactorY = (Ypixels / DesignY)
    ScaleMode = 1  ' twips
    'Exit Sub  ' uncomment to see how Form1 looks without resizing
    
    
    Resize_For_Resolution ScaleFactorX, ScaleFactorY, MyForm
    
'    Debug.Print "Resolucion: " & Str$(Xpixels) & "*" & Str$(Ypixels)
'    Debug.Print "ResolucionTwips: " & Str$(Screen.Width) & "*" & Str$(Screen.Height)
'    Debug.Print "Width: " & MyForm.Width & "* Height: " & MyForm.Height
    MyForm.Top = (Screen.Height / 2 - (MyForm.Height / 2)) / 3
    MyForm.Left = (Screen.Width / 2) - (MyForm.Width / 2)
    MyFormSize.Height = MyForm.Height ' Remember the current size
    MyFormSize.Width = MyForm.Width
End Sub

'Public Sub Ajusta(ByVal frm As Form)
'
'    If Not DoResize Then  ' To avoid infinite loop
'       DoResize = True
'       Exit Sub
'    End If
'
'    RePosForm = False
'
'    ScaleFactorX = frm.Width / MyFormSize.Width   ' How much change?
'    ScaleFactorY = frm.Height / MyFormSize.Height
'    Continua = True
'
'    Resize_For_Resolution ScaleFactorX, ScaleFactorY, frm
'
'End Sub


Public Sub Reajusta(ByVal frm As Object)
    Dim Pantalla As RECT
    Static BrincaFrm As Boolean
    
    If Not BrincaFrm Then
    If frm.WindowState <> vbMinimized Then 'And frm.WindowState <> vbMaximized Then
        If frm.WindowState <> vbMaximized Then
   
            Pantalla = ObtenAreaTrabajo
            'Evalua si el Form es mas grande que las dimensiones actuales de pantalla
            If frm.Height / Screen.TwipsPerPixelY > (Pantalla.Bottom - Pantalla.Top) Then
                BrincaFrm = True
                frm.Height = (Pantalla.Bottom - Pantalla.Top) * Screen.TwipsPerPixelY
            End If
            If frm.Width / Screen.TwipsPerPixelX > (Pantalla.Right - Pantalla.Left) Then
                BrincaFrm = True
                frm.Width = (Pantalla.Right - Pantalla.Left) * Screen.TwipsPerPixelX
            End If
        End If
        frm.AutoRedraw = False
        BrincaFrm = False
        ResizeForm frm
        frm.AutoRedraw = True
    End If
    End If
End Sub

'Workfile:      RS_FORM.BAS
'Created:       07/10/97
'Author:        David Thieme
'Description:   This module provides the code needed to
'               adjust the placement of all controls on
'               a form. There are three public subs.
'               How to use this module:
'                   In a forms Resize event type
'                       ResizeForm Me
'                           - This will resize all controls
'                             on the form to match new form size
'                   You can save a default form size by calling
'                       StoreFormPosition Me
'                   You can restore a form to its original size or
'                   the size that was stored using the StoreFormPosition
'                   sub by calling
'                       RestoreFormPosition Me
'Dependencies:  None
'Issues:        No known problems
'                   Please E-Mail problems to davet@paonline.com

Private Function ActualPos(plLeft As Long) As Long
    If plLeft < 0 Then
        ActualPos = plLeft + 75000
    Else
        ActualPos = plLeft
    End If
End Function

Private Function FindForm(pfrmIn As Form) As Long
    Dim i As Long
    FindForm = -1
    If MaxForm > 0 Then
        For i = 0 To (MaxForm - 1)
            If FormRecord(i).Name = pfrmIn.Name Then
                FindForm = i
                Exit Function
            End If
        Next i
    End If
End Function

Private Function AddForm(pfrmIn As Form) As Long
    Dim FormControl As Control
    Dim i As Long
    ReDim Preserve FormRecord(MaxForm + 1)
    FormRecord(MaxForm).Name = pfrmIn.Name
    FormRecord(MaxForm).Top = pfrmIn.Top
    FormRecord(MaxForm).Left = pfrmIn.Left
    FormRecord(MaxForm).Height = pfrmIn.Height
    FormRecord(MaxForm).Width = pfrmIn.Width
    FormRecord(MaxForm).ScaleHeight = pfrmIn.ScaleHeight
    FormRecord(MaxForm).ScaleWidth = pfrmIn.ScaleWidth
    AddForm = MaxForm
    MaxForm = MaxForm + 1
    For Each FormControl In pfrmIn
        i = FindControl(FormControl, pfrmIn.Name)
        If i < 0 Then
            i = AddControl(FormControl, pfrmIn.Name)
        End If
    Next FormControl
End Function

Private Function FindControl(inControl As Control, inName As String) As Long
    Dim i As Long
    FindControl = -1
    For i = 0 To (MaxControl - 1)
        If ControlRecord(i).Parrent = inName Then
            If ControlRecord(i).Name = inControl.Name Then
                On Error Resume Next
                If ControlRecord(i).Index = inControl.Index Then
                    FindControl = i
                    Exit Function
                End If
                On Error GoTo 0
            End If
        End If
    Next i
End Function

Private Function AddControl(inControl As Control, inName As String) As Long
    ReDim Preserve ControlRecord(MaxControl + 1)
    On Error Resume Next
    ControlRecord(MaxControl).Name = inControl.Name
    ControlRecord(MaxControl).Index = inControl.Index
    ControlRecord(MaxControl).Parrent = inName
    If TypeOf inControl Is Line Then
        ControlRecord(MaxControl).Top = inControl.Y1
        ControlRecord(MaxControl).Left = ActualPos(inControl.X1)
        ControlRecord(MaxControl).Height = inControl.Y2
        ControlRecord(MaxControl).Width = ActualPos(inControl.X2)
    Else
        ControlRecord(MaxControl).Top = inControl.Top
        ControlRecord(MaxControl).Left = ActualPos(inControl.Left)
        ControlRecord(MaxControl).Height = inControl.Height
        ControlRecord(MaxControl).Width = inControl.Width
    End If
    'If TypeOf inControl Is DBList Then
    '    inControl.IntegralHeight = False
    'End If
    On Error GoTo 0
    AddControl = MaxControl
    MaxControl = MaxControl + 1
End Function

Private Function PerWidth(pfrmIn As Form) As Long
    Dim i As Long
   On Error Resume Next
    i = FindForm(pfrmIn)
    If i < 0 Then
        i = AddForm(pfrmIn)
    End If
    
    PerWidth = (pfrmIn.ScaleWidth * 100) \ FormRecord(i).ScaleWidth
End Function

Private Function PerHeight(pfrmIn As Form) As Single
    Dim i As Long
    i = FindForm(pfrmIn)
    If i < 0 Then
        i = AddForm(pfrmIn)
    End If
    PerHeight = (pfrmIn.ScaleHeight * 100) \ FormRecord(i).ScaleHeight
End Function

Private Sub ResizeControl(inControl As Control, pfrmIn As Form)
    Dim i As Long
    Dim yRatio, xRatio, lTop, lLeft, lWidth, lHeight As Long
    yRatio = PerHeight(pfrmIn)
    xRatio = PerWidth(pfrmIn)
    i = FindControl(inControl, pfrmIn.Name)
    On Error GoTo Moveit
    If inControl.Left < 0 Then
        lLeft = CLng(((ControlRecord(i).Left * xRatio) \ 100) - 75000)
    Else
        lLeft = CLng((ControlRecord(i).Left * xRatio) \ 100)
    End If
    lTop = CLng((ControlRecord(i).Top * yRatio) \ 100)
    lWidth = CLng((ControlRecord(i).Width * xRatio) \ 100)
    lHeight = CLng((ControlRecord(i).Height * yRatio) \ 100)
    GoTo Moveit
Moveit:
    On Error GoTo MoveError1
    If TypeName(inControl) = "Line" Then
        If inControl.X1 < 0 Then
            inControl.X1 = CLng(((ControlRecord(i).Left * xRatio) \ 100) - 75000)
        Else
            inControl.X1 = CLng((ControlRecord(i).Left * xRatio) \ 100)
        End If
        inControl.Y1 = CLng((ControlRecord(i).Top * yRatio) \ 100)
        If inControl.X2 < 0 Then
            inControl.X2 = CLng(((ControlRecord(i).Width * xRatio) \ 100) - 75000)
        Else
            inControl.X2 = CLng((ControlRecord(i).Width * xRatio) \ 100)
        End If
        inControl.Y2 = CLng((ControlRecord(i).Height * yRatio) \ 100)
    Else
        If TypeName(inControl) = "Timer" Then
            GoTo subExit
        End If
        If TypeName(inControl) = "Image" Then ' ImageList
            GoTo subExit
        End If
        If TypeName(inControl) = "ImageList" Then
            GoTo subExit
        End If
        
        If TypeName(inControl) = "CommonDialog" Then
          GoTo subExit
        End If
        If TypeName(inControl) = "Timer" Then
            GoTo subExit
        End If
        'If TypeOf inControl Is Winsock Then
        '    GoTo subExit
        'End If
       If TypeName(inControl) = "Menu" Then
        GoTo subExit
       End If
       
       If inControl.Tag <> 1 Then
       
            If TypeName(inControl) = "ComboBox" Then
                inControl.Move lLeft, lTop, lWidth
            Else
                inControl.Move lLeft, lTop, lWidth, lHeight
            End If
        End If
     
    End If
    GoTo subExit
MoveError1:
    On Error GoTo MoveError2
    inControl.Move lLeft, lTop, lWidth
    GoTo subExit
MoveError2:
    On Error GoTo subExit
    inControl.Move lLeft, lTop
subExit:
    On Error GoTo 0
End Sub

Private Sub ResizeForm(pfrmIn As Form)
    Dim FormControl As Control
    Dim isVisible As Boolean
    If pfrmIn.Top < 30000 Then
          
        
        isVisible = pfrmIn.Visible
        'pfrmIn.Visible = False
        For Each FormControl In pfrmIn
            
            ResizeControl FormControl, pfrmIn
        Next FormControl
        'pfrmIn.Visible = isVisible
    End If
End Sub

Public Sub SaveFormPosition(pfrmIn As Object)
    Dim i As Long
    If MaxForm > 0 Then
        For i = 0 To (MaxForm - 1)
            If FormRecord(i).Name = pfrmIn.Name Then
                FormRecord(i).Top = pfrmIn.Top
                FormRecord(i).Left = pfrmIn.Left
                FormRecord(i).Height = pfrmIn.Height
                FormRecord(i).Width = pfrmIn.Width
                Exit Sub
            End If
        Next i
        AddForm (pfrmIn)
    End If
End Sub

Public Sub RestoreFormPosition(pfrmIn As Object)
    Dim i As Long
    If MaxForm > 0 Then
        For i = 0 To (MaxForm - 1)
            If FormRecord(i).Name = pfrmIn.Name Then
                If FormRecord(i).Top < 0 Then
                    pfrmIn.WindowState = 2
                ElseIf FormRecord(i).Top < 30000 Then
                    pfrmIn.WindowState = 0
                    pfrmIn.Move FormRecord(i).Left, FormRecord(i).Top, FormRecord(i).Width, FormRecord(i).Height
                Else
                    pfrmIn.WindowState = 1
                End If
                Exit Sub
            End If
        Next i
    End If
End Sub




