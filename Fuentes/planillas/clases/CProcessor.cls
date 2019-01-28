VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'**************************************************
'* Title       : Chip 8083 FX                     *
'* Programmer  : Deri Dayandri                    *
'* Date        : 10 November 2000                 *
'* Description : Calculate simple math expression *
'**************************************************

Private Const MODULE_NAME = "Chip 8083 FX"

'*** states and operators
Private Const STATE_NONE = 0
Private Const STATE_OPERAND = 1
Private Const STATE_OPERATOR = 2
Private Const STATE_UNARYOP = 3

Private Const UNARY_NEG = "~"

Private Const BINARY_PLUS = "+"
Private Const BINARY_MINUS = "-"
Private Const BINARY_POWER = "^"
Private Const BINARY_MULTIPLE = "*"
Private Const BINARY_DEVIDE = "/"

Private Const PARENTHESIS_OPEN = "("
Private Const PARENTHESIS_CLOSE = ")"

Private Const SPACE = " "


'*** Standard functions
Private Const F_MEMORY = "MEM"
Private Const F_ABSOLOUTE = "ABS"
Private Const F_SQUARE = "SQR"
Private Const F_SQUAREROOT = "SQRT"
Private Const F_INTEGER = "INT"
Private Const F_DECIMAL = "DEC"
Private Const F_LOGARITHM = "LOG"
Private Const F_NATURAL_LOGARITHM = "ALOG"
Private Const F_RADIAN = "RAD"
Private Const F_DEGREE = "DEG"
Private Const F_SINE = "SIN"
Private Const F_COSINE = "COSIN"
Private Const F_TANGENT = "TAN"
Private Const F_ARCTANGENT = "ATAN"

Private Const ListFunctionsWithParameter = ";MEM;ABS;SQR;SQRT;INT;DEC" & _
                                           ";LOG;ALOG;RAD;DEG;SIN;COSIN;TAN;ATAN;"

Private Const F_LASTRESULT = "LRES"
Private Const F_RANDOM = "RND"
Private Const F_PHI = "PHI"

Private Const ListFunctionsWithoutParameter = ";LRES;RND;PHI;"


'*** functions that derived from standard one
Private Const F_SECANT = "SEC"
Private Const F_COSECANT = "COSEC"
Private Const F_COTANGENT = "COTAN"
Private Const F_ARCSINE = "ASIN"
Private Const F_ARCCOSINE = "ACOSIN"
Private Const F_ARCSECANT = "ASEC"
Private Const F_ARCCOSECANT = "ACOSEC"
Private Const F_ARCCOTANGENT = "ACOTAN"

Private Const ListDerivedFunctions1 = ";SEC;COSEC;COTAN" & _
                                      ";ASIN;ACOSIN;ASEC" & _
                                      ";ACOSEC;ACOTAN;"

Private Const F_HYPERBOLIC_SINE = "HSIN"
Private Const F_HYPERBOLIC_COSINE = "HCOSIN"
Private Const F_HYPERBOLIC_TANGENT = "HTAN"
Private Const F_HYPERBOLIC_SECANT = "HSEC"
Private Const F_HYPERBOLIC_COSECANT = "HCOSEC"
Private Const F_HYPERBOLIC_COTANGENT = "HCOTAN"
Private Const F_HYPERBOLIC_ARCSINE = "HASIN"
Private Const F_HYPERBOLIC_ARCCOSINE = "HACOSIN"
Private Const F_HYPERBOLIC_ARCTANGENT = "HATAN"
Private Const F_HYPERBOLIC_ARCSECANT = "HASEC"
Private Const F_HYPERBOLIC_ARCCOSECANT = "HACOSEC"
Private Const F_HYPERBOLIC_ARCCOTANGENT = "HACOTAN"

Private Const ListDerivedFunctions2 = ";HSIN;HCOSIN;HTAN" & _
                                      ";HSEC;HCOSEC;HCOTAN" & _
                                      ";HASIN;HACOSIN;HATAN" & _
                                      ";HASEC;HACOSEC;HACOTAN;"


'*** error code list
Private Const ERROR_SYNTAX_ERROR = 10100
Private Const ERROR_OPERATOR_EXPECTED = 10101
Private Const ERROR_OPERAND_EXPECTED = 10102
Private Const ERROR_PARENTHESIS_NOT_MATCH = 10103
Private Const ERROR_CLOSING_PARENTHESIS_EXPECTED = 10104
Private Const ERROR_INVALID_DECIMAL_POINT = 10105
Private Const ERROR_INVALID_EXPONENT = 10106
Private Const ERROR_OPEN_PARENTHESIS_EXPECTED = 10107
Private Const ERROR_INVALID_MEMORY_INDEX = 10200
Private Const ERROR_FUNCTION_NOT_SUPPORTED = 10300

Private Type TYPE_ERROR
  Code As Integer
  Position As Integer
  Decription As String
End Type

Private vError As TYPE_ERROR


'*** calculation method
Public Enum BaseCalculation
  RadianMode = 0
  DegreeMode = 1
End Enum

Private vBaseCalculation As BaseCalculation

'*** others
Private vLastResult As Double
Private vMemory() As Double
Private vFunctions As Collection
Private vLastFunction As String
Private vConstants As Collection




'**** Public function for Chip 8083 FX

Public Function Calculate(ByVal Expression As String) As Double
'calculate string expression and find the value

Dim Buffer As String

  InfixToPostfix Expression, Buffer
  If vError.Code > 0 Then
    Err.Raise vError.Code, MODULE_NAME, _
              "[" & vError.Code & "] " & _
               vError.Decription
  Else
    vLastResult = DoEvaluate(Buffer)
    Calculate = vLastResult
  End If
End Function

Public Property Let Memory(Index As Integer, vData As Double)
'save something in memory and can be used in Mem function

  If Index < 0 Then
    vError.Code = ERROR_INVALID_MEMORY_INDEX
    vError.Decription = GetErrorDescription()
    Err.Raise vError.Code, MODULE_NAME, _
              "[" & vError.Code & "] " & _
               vError.Decription
  End If
  
  If Index > UBound(vMemory) Then
    ReDim Preserve vMemory(Index)
  End If
  vMemory(Index) = vData
End Property

Public Property Get Memory(Index As Integer) As Double
'get something that store in memory

  If Index > UBound(vMemory) Then Exit Property
  Memory = vMemory(Index)
End Property

Public Sub ReleaseMemory()
'release storable memory

  Erase vMemory
  ReDim vMemory(0)
End Sub

Public Property Let BaseCalculation(vData As BaseCalculation)
'set how to calculate angle, base on radian or base on degree

  vBaseCalculation = vData
End Property

Public Property Get BaseCalculation() As BaseCalculation
'get how to calculate angle

  BaseCalculation = vBaseCalculation
End Property

Public Property Let DeclareFunction(FunctionName As String, Formula As String)
'define additional function

  On Error Resume Next
  vFunctions.Add Formula, UCase(FunctionName)
  If Err <> 0 Then
    vFunctions.Remove FunctionName
    vFunctions.Add Formula, UCase(FunctionName)
  End If
  On Error GoTo 0
  Err.Clear
End Property

Public Property Get DeclareFunction(FunctionName As String) As String
'get formula of additional function

  On Error Resume Next
  DeclareFunction = vFunctions(FunctionName)
  On Error GoTo 0
  Err.Clear
End Property

Public Property Let DeclareConstant(ConstantName As String, Value As Double)
'declare additional constant

  On Error Resume Next
  vConstants.Add Value, UCase(Trim(ConstantName))
  If Err <> 0 Then
    vConstants.Remove ConstantName
    vConstants.Add Value, UCase(ConstantName)
  End If
  On Error GoTo 0
  Err.Clear
End Property

Public Property Get DeclareConstant(ConstantName As String) As Double
'get constant value

  DeclareConstant = 0
  On Error Resume Next
  DeclareConstant = CDbl((vConstants(Trim(ConstantName))))
  On Error GoTo 0
  Err.Clear
End Property

'**** Private function for Chip 8083 FX

Private Sub InfixToPostfix(Expression As String, Buffer As String)
    'check for the expression and convert it to postfix form
    
    Dim i As Integer
    Dim Ch As String
    Dim Token As String
    Dim CurrentState As Integer
    Dim ParenthesisCount As Integer
    Dim HaveDecimal As Boolean
    Dim HaveExponent As Boolean
    Dim StackTokens As CStack

    With vError
        .Code = 0
        .Position = 0
        .Decription = ""
    End With
  
  RemoveSpace Expression
  
  Set StackTokens = New CStack
  
  CurrentState = STATE_NONE
  ParenthesisCount = 0
  i = 1
  Do Until i > Len(Expression)
    Ch = UCase(Mid(Expression, i, 1))
    Select Case Ch
      Case PARENTHESIS_OPEN
        If CurrentState = STATE_OPERAND Then
          vError.Code = ERROR_OPERATOR_EXPECTED
          GoTo ExitFunction
        End If
        
        If CurrentState = STATE_UNARYOP Then
          CurrentState = STATE_OPERATOR
        End If
        
        StackTokens.Push Ch
        ParenthesisCount = ParenthesisCount + 1
        
      Case PARENTHESIS_CLOSE
        If CurrentState <> STATE_OPERAND Then
          vError.Code = ERROR_OPERAND_EXPECTED
          GoTo ExitFunction
        End If
        
        If ParenthesisCount = 0 Then
          vError.Code = ERROR_PARENTHESIS_NOT_MATCH
          GoTo ExitFunction
        End If
        
        Token = StackTokens.Pop
        Do Until Right(Token, 1) = PARENTHESIS_OPEN
          Buffer = Buffer & Token & SPACE
          Token = StackTokens.Pop
        Loop
        
        If Len(Token) <> 1 Then
          Buffer = Buffer & Left(Token, Len(Token) - 1) & SPACE
        End If
        
        ParenthesisCount = ParenthesisCount - 1
          
      Case BINARY_PLUS, _
           BINARY_MINUS, _
           BINARY_MULTIPLE, _
           BINARY_DEVIDE, _
           BINARY_POWER
        If CurrentState = STATE_OPERAND Then
          Do While StackTokens.StackSize > 0
            If GetLevel(StackTokens.GetPopValue) < GetLevel(Ch) Then
              Exit Do
            End If
            Buffer = Buffer & StackTokens.Pop & SPACE
          Loop
          
          StackTokens.Push Ch
          CurrentState = STATE_OPERATOR
        
        ElseIf CurrentState = STATE_UNARYOP Then
          vError.Code = ERROR_OPERAND_EXPECTED
          GoTo ExitFunction
          
        Else
          If Ch = BINARY_MINUS Then
            StackTokens.Push UNARY_NEG
            CurrentState = STATE_UNARYOP
            
          ElseIf Ch = BINARY_PLUS Then
            CurrentState = STATE_UNARYOP
          Else
            vError.Code = ERROR_OPERAND_EXPECTED
            GoTo ExitFunction
          End If
        End If
          
      Case "0" To "9", "."
        If CurrentState = STATE_OPERAND Then
          vError.Code = ERROR_OPERATOR_EXPECTED
          GoTo ExitFunction
        End If
        
        Token = ""
        HaveDecimal = False
        HaveExponent = False
        Do While InStr("0123456789.E", Ch)
          If Ch = "." Then
            If HaveDecimal Then
              vError.Code = ERROR_INVALID_DECIMAL_POINT
              GoTo ExitFunction
            Else
              HaveDecimal = True
            End If
            
          ElseIf Ch = "E" Then
            If HaveExponent Then
              vError.Code = ERROR_INVALID_EXPONENT
              GoTo ExitFunction
            Else
              HaveExponent = True
              Select Case Mid(Expression, i + 1, 1)
                Case BINARY_PLUS, BINARY_MINUS
                  Token = Token & Ch
                  i = i + 1
                  If i > Len(Expression) Then Exit Do
                  Ch = UCase(Mid(Expression, i, 1))
              End Select
            End If
          End If
          Token = Token & Ch
          i = i + 1
          If i > Len(Expression) Then Exit Do
          Ch = UCase(Mid(Expression, i, 1))
        Loop
        i = i - 1
        
        If Token = "." Then
          vError.Code = ERROR_OPERAND_EXPECTED
          GoTo ExitFunction
        End If
        
        If Not IsNumeric(Token) Then
          vError.Code = ERROR_SYNTAX_ERROR
          GoTo ExitFunction
        End If
        
        Buffer = Buffer & Token & SPACE
        CurrentState = STATE_OPERAND
          
      Case "A" To "Z"
        If CurrentState = STATE_OPERAND Then
          vError.Code = ERROR_OPERATOR_EXPECTED
          GoTo ExitFunction
        End If
        
        Token = ""
        Do While InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890_", Ch)
          Token = Token & Ch
          i = i + 1
          If i > Len(Expression) Then Exit Do
          Ch = UCase(Mid(Expression, i, 1))
        Loop
        
        If InStr(ListFunctionsWithParameter & _
                 ListDerivedFunctions1 & _
                 ListDerivedFunctions2, Token) > 0 Then
          If Ch <> PARENTHESIS_OPEN Then
            vError.Code = ERROR_OPEN_PARENTHESIS_EXPECTED
            GoTo ExitFunction
          Else
            StackTokens.Push Token & Ch
            ParenthesisCount = ParenthesisCount + 1
          End If
          
        ElseIf InStr(ListFunctionsWithoutParameter, Token) > 0 Then
          i = i - 1
          Buffer = Buffer & Token & SPACE
          CurrentState = STATE_OPERAND
          
        ElseIf FunctionExist(Token) Then
          If Ch <> PARENTHESIS_OPEN Then
            vError.Code = ERROR_OPEN_PARENTHESIS_EXPECTED
            GoTo ExitFunction
          Else
            StackTokens.Push Token & Ch
            ParenthesisCount = ParenthesisCount + 1
          End If
          
        Else
          If Ch = PARENTHESIS_OPEN Then
            vLastFunction = Token
            vError.Code = ERROR_FUNCTION_NOT_SUPPORTED
            GoTo ExitFunction
          Else
            i = i - 1
            Buffer = Buffer & DeclareConstant(Token) & SPACE
            CurrentState = STATE_OPERAND
          End If
        End If
    End Select
    i = i + 1
  Loop
  
  If CurrentState = STATE_OPERATOR Or _
     CurrentState = STATE_UNARYOP Then
    vError.Code = ERROR_OPERAND_EXPECTED
    GoTo ExitFunction
  End If
  
  If ParenthesisCount > 0 Then
    vError.Code = ERROR_CLOSING_PARENTHESIS_EXPECTED
    GoTo ExitFunction
  End If
  
  Do Until StackTokens.StackSize = 0
    Buffer = Buffer & StackTokens.Pop & SPACE
  Loop
  
ExitFunction:
  vError.Position = i
  GetErrorDescription
  Set StackTokens = Nothing
  Exit Sub
End Sub

Private Function GetLevel(Operator As String) As Integer
'get the current level of operator

  Select Case Operator
    Case BINARY_PLUS, BINARY_MINUS
      GetLevel = 1
    Case BINARY_MULTIPLE, BINARY_DEVIDE
      GetLevel = 2
    Case BINARY_POWER
      GetLevel = 3
    Case UNARY_NEG
      GetLevel = 4
    Case Else
      GetLevel = 0
  End Select
End Function

Private Function DoEvaluate(Expression As String) As Double
'get the value of the expression
'the expresion must be in postfix form

Dim i As Integer
Dim J As Integer
Dim Token As String
Dim X As Variant
Dim Y As Variant
Dim StackTokens As CStack
Dim NewExpression As String

  Set StackTokens = New CStack
  i = 1
  J = InStr(Expression, SPACE)
  
  On Error GoTo Error_Evaluate
  
  Do Until J = 0
    Token = Mid(Expression, i, J - i)
    If IsNumeric(Token) Then
      StackTokens.Push CDbl(Token)
      
    Else
      Select Case Token
        '*** standard operations
        Case UNARY_NEG
          StackTokens.Push -StackTokens.Pop
            
        Case BINARY_DEVIDE
          X = StackTokens.Pop
          Y = StackTokens.Pop
          
          StackTokens.Push Y / X
            
        Case BINARY_MULTIPLE
          StackTokens.Push StackTokens.Pop * StackTokens.Pop
            
        Case BINARY_PLUS
          StackTokens.Push StackTokens.Pop + StackTokens.Pop
          
        Case BINARY_MINUS
          X = StackTokens.Pop
          Y = StackTokens.Pop
          StackTokens.Push Y - X
          
        Case BINARY_POWER
          X = StackTokens.Pop
          Y = StackTokens.Pop
          StackTokens.Push Y ^ X
          
          
        '*** standard functions
        Case F_MEMORY
          X = StackTokens.Pop
          X = Abs(Int(X))
          If X > UBound(vMemory) Then
            StackTokens.Push 0
          Else
            StackTokens.Push vMemory(X)
          End If
          
        Case F_ABSOLOUTE
          X = StackTokens.Pop
          StackTokens.Push Abs(X)
          
        Case F_SQUARE
          X = StackTokens.Pop
          StackTokens.Push X * X
        
        Case F_SQUAREROOT
          X = StackTokens.Pop
          StackTokens.Push Sqr(X)
          
        Case F_INTEGER
          X = StackTokens.Pop
          StackTokens.Push Int(X)
          
        Case F_DECIMAL
          X = StackTokens.Pop
          If X > 0 Then
            StackTokens.Push X - Int(X)
          Else
            StackTokens.Push -(Abs(X) - Int(Abs(X)))
          End If
          
        Case F_LOGARITHM
          X = StackTokens.Pop
          StackTokens.Push Log(X)
          
        Case F_NATURAL_LOGARITHM
          X = StackTokens.Pop
          StackTokens.Push Exp(X)
          
        Case F_RADIAN
          X = StackTokens.Pop
          StackTokens.Push X * Atn(1) / 45
          
        Case F_DEGREE
          X = StackTokens.Pop
          StackTokens.Push X * 45 / Atn(1)
        
        Case F_SINE
          X = StackTokens.Pop
          If vBaseCalculation = DegreeMode Then X = X * Atn(1) / 45
          StackTokens.Push Sin(X)
        
        Case F_COSINE
          X = StackTokens.Pop
          If vBaseCalculation = DegreeMode Then X = X * Atn(1) / 45
          StackTokens.Push Cos(X)
        
        Case F_TANGENT
          X = StackTokens.Pop
          If vBaseCalculation = DegreeMode Then X = X * Atn(1) / 45
          StackTokens.Push Tan(X)
        
        Case F_ARCTANGENT
          X = StackTokens.Pop
          X = Atn(X)
          If vBaseCalculation = DegreeMode Then X = X * 45 / Atn(1)
          StackTokens.Push X
          
        Case F_LASTRESULT
          StackTokens.Push vLastResult
          
        Case F_PHI
          StackTokens.Push Atn(1) * 4
          
        Case F_RANDOM
          StackTokens.Push Rnd
          
        
        '*** derived functions
        Case F_SECANT
          X = StackTokens.Pop
          If vBaseCalculation = DegreeMode Then X = X * Atn(1) / 45
          StackTokens.Push 1 / Cos(X)
          
        Case F_COSECANT
          X = StackTokens.Pop
          If vBaseCalculation = DegreeMode Then X = X * Atn(1) / 45
          StackTokens.Push 1 / Sin(X)
        
        Case F_COTANGENT
          X = StackTokens.Pop
          If vBaseCalculation = DegreeMode Then X = X * Atn(1) / 45
          StackTokens.Push 1 / Tan(X)
        
        Case F_ARCSINE
          X = StackTokens.Pop
          X = Atn(X / Sqr(-X * X + 1))
          If vBaseCalculation = DegreeMode Then X = X * 45 / Atn(1)
          StackTokens.Push X
        
        Case F_ARCCOSINE
          X = StackTokens.Pop
          X = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
          If vBaseCalculation = DegreeMode Then X = X * 45 / Atn(1)
          StackTokens.Push X
        
        Case F_ARCSECANT
          X = StackTokens.Pop
          X = Atn(X / Sqr(X * X - 1)) + Sgn((X) - 1) * (2 * Atn(1))
          If vBaseCalculation = DegreeMode Then X = X * 45 / Atn(1)
          StackTokens.Push X
        
        Case F_ARCCOSECANT
          X = StackTokens.Pop
          X = Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1))
          If vBaseCalculation = DegreeMode Then X = X * 45 / Atn(1)
          StackTokens.Push X
        
        Case F_ARCCOTANGENT
          X = StackTokens.Pop
          X = Atn(X) + 2 * Atn(1)
          If vBaseCalculation = DegreeMode Then X = X * 45 / Atn(1)
          StackTokens.Push X
          
        Case F_HYPERBOLIC_SINE
          X = StackTokens.Pop
          StackTokens.Push (Exp(X) - Exp(-X)) / 2
        
        Case F_HYPERBOLIC_COSINE
          X = StackTokens.Pop
          StackTokens.Push (Exp(X) + Exp(-X)) / 2
        
        Case F_HYPERBOLIC_TANGENT
          X = StackTokens.Pop
          StackTokens.Push (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
        
        Case F_HYPERBOLIC_SECANT
          X = StackTokens.Pop
          StackTokens.Push 2 / (Exp(X) + Exp(-X))
        
        Case F_HYPERBOLIC_COSECANT
          X = StackTokens.Pop
          StackTokens.Push 2 / (Exp(X) - Exp(-X))
        
        Case F_HYPERBOLIC_COTANGENT
          X = StackTokens.Pop
          StackTokens.Push (Exp(X) + Exp(-X)) / (Exp(X) - Exp(-X))
        
        Case F_HYPERBOLIC_ARCSINE
          X = StackTokens.Pop
          StackTokens.Push Log(X + Sqr(X * X + 1))
        
        Case F_HYPERBOLIC_ARCCOSINE
          X = StackTokens.Pop
          StackTokens.Push Log(X + Sqr(X * X - 1))
        
        Case F_HYPERBOLIC_ARCTANGENT
          X = StackTokens.Pop
          StackTokens.Push Log((1 + X) / (1 - X)) / 2
        
        Case F_HYPERBOLIC_ARCSECANT
          X = StackTokens.Pop
          StackTokens.Push Log((Sqr(-X * X + 1) + 1) / X)
        
        Case F_HYPERBOLIC_ARCCOSECANT
          X = StackTokens.Pop
          StackTokens.Push Log((Sgn(X) * Sqr(X * X + 1) + 1) / X)
        
        Case F_HYPERBOLIC_ARCCOTANGENT
          X = StackTokens.Pop
          StackTokens.Push Log((X + 1) / (X - 1)) / 2
      
        'additional functions
        Case Else
          X = StackTokens.Pop
          NewExpression = GetFunction(Token, CDbl(X))
          StackTokens.Push Calculate(NewExpression)
      End Select
    End If
    
    i = J + 1
    J = InStr(i, Expression, SPACE)
  Loop
  
  If StackTokens.StackSize > 0 Then
    DoEvaluate = StackTokens.Pop
  
  Else
    DoEvaluate = 0
  End If
  Exit Function
  
Error_Evaluate:
  vError.Code = Err.Number
  vError.Decription = Err.Description
  Err.Clear
  Set StackTokens = Nothing
  
  On Error GoTo 0
'  Err.Raise vError.Code, MODULE_NAME, _
            "[" & vError.Code & "] " & vError.Decription
            
End Function

Private Function GetErrorDescription() As String
'get the description of the current error code

  With vError
    Select Case .Code
      Case ERROR_SYNTAX_ERROR
        .Decription = "Syntax error at column " & .Position
      Case ERROR_OPERATOR_EXPECTED
        .Decription = "Expected operator at column " & .Position
      Case ERROR_OPERAND_EXPECTED
        .Decription = "Expected operand at column " & .Position
      Case ERROR_PARENTHESIS_NOT_MATCH
        .Decription = "The number of open parenthesis is not match with the number of close parenthesis."
      Case ERROR_CLOSING_PARENTHESIS_EXPECTED
        .Decription = "Expected closing parenthesis at column " & .Position
      Case ERROR_INVALID_DECIMAL_POINT
        .Decription = "Invalid use of point at column " & .Position
      Case ERROR_INVALID_EXPONENT
        .Decription = "Invalid use of exponent"
      Case ERROR_OPEN_PARENTHESIS_EXPECTED
        .Decription = "Expected openparenthesis at column " & .Position
      Case ERROR_INVALID_MEMORY_INDEX
        .Decription = "Invalid memory index"
      Case ERROR_FUNCTION_NOT_SUPPORTED
        .Decription = "Function " & vLastFunction & " is not defined yet"
      Case Else
        .Decription = ""
    End Select
  End With
End Function

Private Sub RemoveSpace(Expression As String)
'Remove all of the space in expression

Dim i As Integer
Dim NewExpression As String

  For i = 1 To Len(Expression)
    If Mid(Expression, i, 1) <> SPACE Then
      NewExpression = NewExpression & Mid(Expression, i, 1)
    End If
  Next i
  Expression = NewExpression
End Sub

Private Function FunctionExist(FunctionName As String) As Boolean
    'check if functionname is defined or not
    
    Dim strTemp As String
    
    On Error Resume Next
'    strTemp = vFunctions(UCase(FunctionName))
'    FunctionExist = (Err = 0)
    On Error GoTo 0
    Err.Clear
End Function

Private Function GetFunction(FunctionName As String, Parameter As Single) As String
'get the defined function expression and
'change the parameter with the real value

Dim i As Integer
Dim RealExpression As String
Dim Expression As String
Dim LastCh As String
Dim NextCh As String
Dim ActCh As String
Dim CheckString As String

  CheckString = PARENTHESIS_OPEN & PARENTHESIS_CLOSE & _
                BINARY_PLUS & BINARY_MINUS & _
                BINARY_MULTIPLE & BINARY_DEVIDE & _
                BINARY_POWER
  
  Expression = vFunctions(FunctionName)
  RemoveSpace Expression
  RealExpression = ""
  For i = 1 To Len(Expression)
    If i = 1 Then LastCh = PARENTHESIS_OPEN
    If i = Len(Expression) Then
      NextCh = PARENTHESIS_CLOSE
    Else
      NextCh = Mid(Expression, i + 1, 1)
    End If
    ActCh = Mid(Expression, i, 1)
    If ActCh = "X" Then
      If InStr(CheckString, LastCh) > 0 And _
         InStr(CheckString, NextCh) > 0 Then
        RealExpression = RealExpression & Parameter
      Else
        RealExpression = RealExpression & ActCh
      End If
    
    Else
      RealExpression = RealExpression & ActCh
    End If
    LastCh = ActCh
  Next i
  
  GetFunction = RealExpression
End Function


Private Sub Class_Initialize()
  Randomize
  ReDim vMemory(0)
  Set vFunctions = New Collection
  Set vConstants = New Collection
  vBaseCalculation = RadianMode
  With vError
    .Code = 0
    .Decription = ""
    .Position = 0
  End With
End Sub

Private Sub Class_Terminate()
  Erase vMemory
  Set vFunctions = Nothing
  Set vConstants = Nothing
End Sub

