VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFeetInches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' CLASS MODULE: cFeetInches
'
' DESCRIPTION:  Class for calculations dealing with feet and inches
'
' REQUIREMENTS: cRebar class module
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/21/2016   Created
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' CLASS DECLARATIONS
'------------------------------------------------------------------------------

Public Enum dimMode
    feet
    inches
End Enum

Public Enum OutputPrecision
    wholeInch = 1
    halfInch = 2
    quarterInch = 4
    eigthInch = 8
    sixteenthInch = 16
End Enum

Private pOutputMode     As dimMode
Private pInputNumber    As Double
Private pFeet           As Long
Private pInches         As Long
Private pInchFraction   As Double
Private pPrecision      As OutputPrecision ' current precision of output
Private pOutExpression  As String

Private Const FTINCHCONVERSION As Double = 12

'------------------------------------------------------------------------------
' CLASS EVENTS
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
    pOutputMode = feet
    pInputNumber = 0#
    pOutExpression = vbNullString
    pFeet = 0
    pInches = 0#
    pPrecision = sixteenthInch
End Sub

'------------------------------------------------------------------------------
' CLASS PROPERTIES
'------------------------------------------------------------------------------
Public Property Get OutputMode() As dimMode
    OutputMode = pOutputMode
End Property

Public Property Let OutputMode(ByVal setMode As dimMode)
    pOutputMode = setMode
End Property

Public Property Get Precision() As OutputPrecision
    Precision = pPrecision
End Property

Public Property Let Precision(ByVal outPrecision As OutputPrecision)
    pPrecision = outPrecision
End Property

Public Property Get dimension() As String
    
    FormatOutputString

    If pOutExpression = vbNullString Then
        dimension = vbNullString
    Else
        dimension = pOutExpression
    End If
End Property

'------------------------------------------------------------------------------
' CLASS METHODS
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DESCRIPTION:  Set the dimension
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/21/2016   Created
'
'------------------------------------------------------------------------------
Public Sub SetDimension(ByVal dimension As Double, _
                        Optional ByVal mode As dimMode = dimMode.feet)
                        
    If Not IsNumeric(dimension) Then ' exit if not a number
        Exit Sub
    End If
    
    pInputNumber = dimension

    If pOutputMode = mode Then
        ParseDecimalFeet number:=dimension
    Else
        ParseDecimalInch number:=dimension
    End If

End Sub

'------------------------------------------------------------------------------
' DESCRIPTION:  Break a decimal foot value into feet and inch variables
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/21/2016   Created
' Tom Gerard    6/22/2016   Added simple error handling
'------------------------------------------------------------------------------
Private Sub ParseDecimalFeet(ByVal number As Double)

    On Error GoTo ErrorHandler

    If CLng(number) = number Then
        pFeet = number
        pInches = 0
        pInchFraction = 0#
    Else
        Dim foot As Long
        foot = Fix(number)
        
        Dim inch As Double
        inch = (Abs(number) - Abs(foot)) * FTINCHCONVERSION
        
        Dim wholeInch As Long
        wholeInch = Fix(inch)
        
        Dim fraction As Double
        fraction = inch - wholeInch
        
        ' check the fraction and adjust variables
        CheckFraction feet:=foot, inch:=wholeInch, fractionalInch:=fraction
        
        pFeet = foot
        pInches = wholeInch
        pInchFraction = fraction
    End If
    
SubExit:
    Exit Sub
    
ErrorHandler:
    ' Open a message box with the error description
    MsgBox "Error Number:  " & Err.number & vbCrLf & Err.Description & vbCrLf & _
        vbCrLf & "Possible Cause" & vbCrLf & "Input number is too large", vbExclamation, _
        "An Error Has Occured"

    Resume SubExit
    
End Sub

'------------------------------------------------------------------------------
' DESCRIPTION:  Break a decimal inch value into feet and inch variables
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/21/2016   Created
'
'------------------------------------------------------------------------------
Private Sub ParseDecimalInch(ByVal number As Double)

    On Error GoTo ErrorHandler

    Dim foot As Long
    foot = Fix(number / FTINCHCONVERSION)
    
    Dim inch As Double
    inch = (Abs(number / FTINCHCONVERSION) - Abs(foot)) * FTINCHCONVERSION
    
    Dim wholeInch As Long
    wholeInch = Fix(inch)
    
    Dim fraction As Double
    fraction = inch - wholeInch
    
    ' check the fraction and adjust variables
    CheckFraction feet:=foot, inch:=wholeInch, fractionalInch:=fraction
    
    pFeet = foot
    pInches = wholeInch
    pInchFraction = fraction
    
SubExit:
    Exit Sub
    
ErrorHandler:
    ' Open a message box with the error description
    MsgBox "Error Number:  " & Err.number & vbCrLf & Err.Description & vbCrLf & _
        vbCrLf & "Possible Cause" & vbCrLf & "Input number is too large", vbExclamation, _
        "An Error Has Occured"

    Resume SubExit

End Sub

'------------------------------------------------------------------------------
' DESCRIPTION:  Format the output string
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/21/2016   Created
'
'------------------------------------------------------------------------------
Private Sub FormatOutputString()
' add inch only output
    If pFeet = 0 And pInches = 0 And pInchFraction = 0 Then
        Exit Sub
    End If

    Dim output As String
    Dim top As Integer
    Dim bottom As Integer
    
    If pOutputMode = dimMode.feet Then 'Output in feet #'-#  #/#"
    
        If pInchFraction > 0 Then
        
            top = Round(pInchFraction * Me.Precision)
            bottom = Me.Precision
            
            ReduceFraction numerator:=top, denominator:=bottom
            
            output = pFeet & "'-" & pInches & "  " & top & "/" & _
                bottom & """"
        Else
            output = pFeet & "'-" & pInches & """"
        End If
        
    Else ' Output in inches #  #/#"
    
        If pInchFraction > 0 Then
        
            top = Round(pInchFraction * Me.Precision)
            bottom = Me.Precision
            
            ReduceFraction numerator:=top, denominator:=bottom
            
            output = pFeet * FTINCHCONVERSION + pInches & "  " & top & "/" & _
                bottom & """"
        Else
            output = pFeet * FTINCHCONVERSION + pInches & """"
        End If
    
    End If
    
    pOutExpression = output
    
End Sub

'------------------------------------------------------------------------------
' DESCRIPTION:  Reduce fractional inch by 2 recursively
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/21/2016   Created
' Tom Gerard    6/22/2016   Added simple error handling
'------------------------------------------------------------------------------
Private Sub ReduceFraction(ByRef numerator As Integer, _
    ByRef denominator As Integer)
    
    On Error GoTo ErrorHandler
    
    If numerator Mod 2 <> 0 Then
        ' exit when no longer divisible by 2
        Exit Sub
    Else
        numerator = numerator / 2
        denominator = denominator / 2
        ReduceFraction numerator, denominator
    End If
    
SubExit:
    Exit Sub
    
ErrorHandler:
    ' Open a message box with the error description
    MsgBox "Error Number:  " & Err.number & vbCrLf & Err.Description & vbCrLf & _
        vbCrLf & "Source: cFeetInches.ReduceFraction", vbExclamation, _
        "An Error Has Occured"

    Resume SubExit
    
End Sub

'------------------------------------------------------------------------------
' DESCRIPTION:  Check fractional inch number and adjust other dimensions if
'               rounded fraction = 1 or inches > 12
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/22/2016   Created
'
'------------------------------------------------------------------------------
Private Sub CheckFraction(ByRef feet As Long, _
                          ByRef inch As Long, _
                          ByRef fractionalInch As Double)
                          
    Dim checkNum As Double
    checkNum = Round(fractionalInch * Me.Precision)
    
    ' adjust numbers if numerator = demonimator
    If (checkNum / Me.Precision) = 1# Then
        
        'adjust inches
        fractionalInch = 0#
        inch = inch + 1
        
        ' adjust feet
        If inch = 12 Then
            inch = 0
            feet = feet + 1
        End If
        
    End If

End Sub

