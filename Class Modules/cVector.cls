VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
' CLASS MODULE: cVector
'
' DESCRIPTION:  Vector class for linear algebra calculations
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/30/2016   Created
'
'----------------------------------------------------------------------------------------
Option Explicit

'----------------------------------------------------------------------------------------
' CLASS DECLARATIONS
'----------------------------------------------------------------------------------------

' Error codes
Private Enum VecErr
    NoAdd = vbObjectError + 522
    NoSubtract
    NoDot
    NoCross
    Range
    NegativeSize
End Enum

' Error constants
Private Const VectorErrorTitle = "Vector Object Error"

Private pSize As Long
Private pVector() As Double

'----------------------------------------------------------------------------------------
' CLASS EVENTS
'----------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    pSize = 0
    ReDim pVector(0)
End Sub ' Class_Initialize

Private Sub Class_Terminate()
    KillObject
End Sub ' Class_Terminate

'----------------------------------------------------------------------------------------
' CLASS PROPERTIES
'----------------------------------------------------------------------------------------
Public Property Get Value(ByVal index As Long) As Double

    On Error GoTo ErrorHandler
    
    Const propName As String = "cVector.Value"
    
    ' check for range error
    If index > UBound(pVector) Then
    
        Err.Raise number:=VecErr.Range, _
                  source:=propName, _
                  description:="Index is out of range!"
    
    End If
    
    Value = pVector(index)
    
PropExit:
    Exit Property
    
ErrorHandler:
    
    Select Case Err.number
    
        Case VecErr.Range
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Error Source : " & Err.source, _
                   Buttons:=vbCritical, _
                   Title:=VectorErrorTitle
        
        Case Else
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number, _
                   Buttons:=vbCritical
    
    End Select
    
    Value = 0
    
    Resume PropExit
    
End Property ' Value(Get)

Public Property Let Value(ByVal index As Long, ByVal setValue As Double)

    On Error GoTo ErrorHandler
    
    Const propName As String = "cVector.Value"
    
    ' check for range error
    If index > UBound(pVector) Then
    
        Err.Raise number:=VecErr.Range, _
                  source:=propName, _
                  description:="Index is out of range!"
    
    End If
    
    pVector(index) = setValue
    
PropExit:
    Exit Property
    
ErrorHandler:
    
    Select Case Err.number
    
        Case VecErr.Range
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Error Source : " & Err.source, _
                   Buttons:=vbCritical, _
                   Title:=VectorErrorTitle
        
        Case Else
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number, _
                   Buttons:=vbCritical
    
    End Select
    
    Resume PropExit

End Property ' Value(Let)

Public Property Get Size() As Long
    Size = pSize
End Property ' Size(Get)

'----------------------------------------------------------------------------------------
' CLASS PUBLIC METHODS
'----------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Sets the size of the vector
'
' INPUTS:       setSize - the size of the vector
'
'               bPreserve - option to save the current contents of the vector
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/30/2016   Created
'
'----------------------------------------------------------------------------------------
Public Sub SetSize(ByVal newSize As Long, Optional ByVal bPreserve As Boolean = False)

    On Error GoTo ErrorHandler
    
    Const subName As String = "cVector.SetSize"
    
    If newSize < 0 Then
    
        Err.Raise number:=VecErr.NegativeSize, _
                  source:=subName, _
                  description:="Vector size cannot be negative"
    
    End If
    
    If bPreserve = True Then
        ReDim Preserve pVector(newSize - 1)
    Else
        ReDim pVector(newSize - 1)
    End If
    
    pSize = newSize
    
SubExit:
    Exit Sub
    
ErrorHandler:
    
    Select Case Err.number
    
        Case VecErr.NegativeSize
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Error Source : " & Err.source, _
                   Buttons:=vbCritical, _
                   Title:=VectorErrorTitle
        
        Case Else
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number, _
                   Buttons:=vbCritical
    
    End Select
    
    Resume SubExit

End Sub ' SetSize

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Populates the entries of the vector from an array of values in the
'               spreadsheet of unkown size
'
' INPUTS:       startRange - first (left or top) cell in vector
'
'               orientation - select horizontal or vertically oriented vector
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Public Sub DynamicVectorFromSheet(ByVal startRange As Range, _
                                  Optional ByVal vertical As Boolean = True)

    On Error GoTo ErrorHandler
    
    Const subName As String = "cVector.DynamicVectorFromSheet"
    
    Dim vecSize As Long
    Dim index As Long
    
    If vertical = True Then
    
        ' read rows from start of range to first empty cell encountered
        vecSize = startRange.CurrentRegion.Rows.Count
        
        pSize = vecSize
        ReDim pVector(vecSize - 1)
        
        For index = 0 To vecSize - 1
            pVector(index) = startRange(index + 1).Value
        Next index
        
    Else ' horizontal
    
        ' read columns from start of range to first empty cell encountered
        vecSize = startRange.CurrentRegion.Columns.Count
        
        pSize = vecSize
        ReDim pVector(vecSize - 1)
        
        For index = 0 To vecSize - 1
            pVector(index) = startRange(1, index + 1).Value
        Next index
    
    End If
    
    
SubExit:
    Exit Sub
    
ErrorHandler:

    MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number, _
                   Buttons:=vbCritical
                   
    Resume SubExit

End Sub ' DynamicVectorFromSheet

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Populates the entries of the vector from an array of values in the
'               spreadsheet of unkown size
'
' INPUTS:       startRange - first (left or top) cell in vector
'
'               orientation - select horizontal or vertically oriented vector
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Public Sub StaticVectorFromSheet(ByVal vectorRange As Range, _
                                 Optional ByVal vertical As Boolean = True)


    On Error GoTo ErrorHandler
    
    Const subName As String = "cVector.StaticVectorFromSheet"
    
    Dim vecSize As Long
    Dim index As Long
    
    If vertical = True Then
    
        vecSize = vectorRange.Rows.Count
        
        pSize = vecSize
        ReDim pVector(vecSize - 1)
        
        For index = 0 To vecSize - 1
            pVector(index) = vectorRange(index + 1).Value
        Next index
    
    Else
    
        vecSize = vectorRange.Columns.Count
        
        pSize = vecSize
        ReDim pVector(vecSize - 1)
        
        For index = 0 To vecSize - 1
            pVector(index) = vectorRange(1, index + 1).Value
        Next index
    
    End If

SubExit:
    Exit Sub
    
ErrorHandler:

    MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number, _
                   Buttons:=vbCritical
                   
    Resume SubExit
    
End Sub ' StaticVectorFromSheet

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Prints vector to Immediate Window
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/30/2016   Created
'
'----------------------------------------------------------------------------------------
Public Sub PrintToConsole()

    On Error GoTo ErrorHandler

    Dim index As Long
    
    For index = 0 To pSize - 1
        Debug.Print pVector(index)
    Next index
    
    Debug.Print

SubExit:
    Exit Sub
    
ErrorHandler:

    MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                   "Error Number : " & Err.number, _
           Buttons:=vbCritical
    
    Resume SubExit
    
End Sub ' PrintToConsole

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Outputs vector to spreadsheet
'
' INPUT:        startCell - location of first cell to print matrix to
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Public Sub OutputToSpreadsheet(ByVal startCell As Range, _
                               Optional ByVal vertical As Boolean = True)

    On Error GoTo ErrorHandler
    
    If vertical = False Then
        startCell.Resize(1, pSize).Value2 = pVector
    Else
        startCell.Resize(pSize, 1).Value2 = WorksheetFunction.Transpose(pVector)
    End If
    
SubExit:
    Exit Sub
    
ErrorHandler:

    MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                   "Error Number : " & Err.number, _
           Buttons:=vbCritical
    
    Resume SubExit
    
End Sub ' OutputToSpreadsheet

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Swaps two values in vector
'
' INPUTS:       index1 - first item to swap
'               index2 - second item to swap
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    7/1/2016   Created
'
'----------------------------------------------------------------------------------------
Public Sub SwapValues(ByVal index1 As Long, ByVal index2 As Long)

    On Error GoTo ErrorHandler
    
    Const subName As String = "cVector.SwapValues"
    
    If index1 > UBound(pVector) Or index2 > UBound(pVector) Then
    
        Err.Raise number:=VecErr.Range, _
                  source:=subName, _
                  description:="Index is out of range!"
    
    End If
    
    Dim temp As Double
    
    temp = pVector(index1)
    pVector(index1) = pVector(index2)
    pVector(index2) = temp
    
SubExit:
    Exit Sub
    
ErrorHandler:
    
    Select Case Err.number
    
        Case VecErr.Range
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Error Source : " & Err.source, _
                   Buttons:=vbCritical, _
                   Title:=VectorErrorTitle
        
        Case Else
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number, _
                   Buttons:=vbCritical
    
    End Select
    
    Resume SubExit

End Sub ' SwapValues

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Creates a copy of the current vector
'
' OUTPUT:       cVector object
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/30/2016   Created
'
'----------------------------------------------------------------------------------------
Public Function Copy() As cVector

    On Error GoTo ErrorHandler
    
    Dim result As cVector
    Set result = New cVector
    
    result.SetSize newSize:=pSize, bPreserve:=False
    
    Dim index As Long
    
    For index = 0 To pSize - 1
        result.Value(index) = pVector(index)
    Next index
    
    Set Copy = result
    
FunctionExit:
    Exit Function
    
ErrorHandler:
    
    MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                   "Error Number : " & Err.number, _
           Buttons:=vbCritical
           
    Resume FunctionExit

End Function ' Copy

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Adds two vectors
'
' INPUT:        otherVector - vector to add to current vector
'
' OUTPUT:       cVector object
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/30/2016   Created
'
'----------------------------------------------------------------------------------------
Public Function Add(ByVal otherVector As cVector) As cVector

    On Error GoTo ErrorHandler
    
    Const funcName As String = "cVector.Add"
    
    If pSize <> otherVector.Size Then
    
        Err.Raise number:=VecErr.NegativeSize, _
                  source:=funcName, _
                  description:="Vectors not compatible for addition!"
    
    End If
    
    Dim result As cVector
    Dim index As Long
    
    Set result = New cVector
    result.SetSize newSize:=pSize
    
    For index = 0 To pSize - 1
        result.Value(index) = pVector(index) + otherVector.Value(index)
    Next index
    
    Set Add = result
    
FunctionExit:
    Exit Function
    
ErrorHandler:

    Select Case Err.number
    
        Case VecErr.NoAdd
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Error Source : " & Err.source, _
                   Buttons:=vbCritical, _
                   Title:=VectorErrorTitle
        
        Case Else
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number, _
                   Buttons:=vbCritical
    
    End Select
    
    Resume FunctionExit

End Function ' Add

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Subtracts two vectors
'
' INPUT:        otherVector - vector to subtract from current vector
'
' OUTPUT:       cVector object
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/30/2016   Created
'
'----------------------------------------------------------------------------------------
Public Function Subtract(ByVal otherVector As cVector) As cVector

    On Error GoTo ErrorHandler
    
    Const funcName As String = "cVector.Subtract"
    
    If pSize <> otherVector.Size Then
    
        Err.Raise number:=VecErr.NegativeSize, _
                  source:=funcName, _
                  description:="Vectors not compatible for subtraction!"
    
    End If
    
    Dim result As cVector
    Dim index As Long
    
    Set result = New cVector
    result.SetSize newSize:=pSize
    
    For index = 0 To pSize - 1
        result.Value(index) = pVector(index) - otherVector.Value(index)
    Next index
    
    Set Subtract = result
    
FunctionExit:
    Exit Function
    
ErrorHandler:

    Select Case Err.number
    
        Case VecErr.NoSubtract
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Error Source : " & Err.source, _
                   Buttons:=vbCritical, _
                   Title:=VectorErrorTitle
        
        Case Else
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number, _
                   Buttons:=vbCritical
    
    End Select
    
    Resume FunctionExit

End Function ' Subtract

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Dot product of two vectors
'
' INPUT:        otherVector - vector to take dot product with
'
' OUTPUT:       scalar double
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/30/2016   Created
'
'----------------------------------------------------------------------------------------
Public Function Dot(ByVal otherVector As cVector) As Double

    On Error GoTo ErrorHandler
    
    Const funcName As String = "cVector.Dot"
    
    If pSize <> otherVector.Size Then
    
        Err.Raise number:=VecErr.NoDot, _
                  source:=funcName, _
                  description:="Vectors not compatible for dot product!"
    
    End If
    
    Dim result As Double
    result = 0#
    
    Dim index As Long
    
    For index = 0 To pSize - 1
        result = result + pVector(index) * otherVector.Value(index)
    Next index
    
    Dot = result
    
FunctionExit:
    Exit Function
    
ErrorHandler:

    Select Case Err.number
    
        Case VecErr.NoDot
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Error Source : " & Err.source, _
                   Buttons:=vbCritical, _
                   Title:=VectorErrorTitle
        
        Case Else
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number, _
                   Buttons:=vbCritical
    
    End Select
    
    Resume FunctionExit

End Function ' Dot

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Dot product of two vectors
'
' INPUT:        otherVector - vector to take dot product with
'
' OUTPUT:       scalar double
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/30/2016   Created
'
'----------------------------------------------------------------------------------------
Public Function Cross(ByVal otherVector As cVector) As cVector

    On Error GoTo ErrorHandler
    
    Const funcName As String = "cVector.Cross"
    
    ' cross product only defined for vectors of size 3
    If pSize <> 3 And pSize <> otherVector.Size Then
    
        Err.Raise number:=VecErr.NoCross, _
                  source:=funcName, _
                  description:="Vectors not compatible for dot product!"
    
    End If
    
    Dim result As cVector
    Set result = New cVector
    
    result.SetSize newSize:=3, bPreserve:=False
    
    With result
        .Value(0) = pVector(1) * otherVector.Value(2) - pVector(2) * otherVector.Value(1)
        .Value(1) = pVector(2) * otherVector.Value(0) - pVector(0) * otherVector.Value(2)
        .Value(2) = pVector(0) * otherVector.Value(1) - pVector(1) * otherVector.Value(0)
    End With
    
    Set Cross = result
    
FunctionExit:
    Exit Function
    
ErrorHandler:

    Select Case Err.number
    
        Case VecErr.NoCross
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Error Source : " & Err.source, _
                   Buttons:=vbCritical, _
                   Title:=VectorErrorTitle
        
        Case Else
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number, _
                   Buttons:=vbCritical
    
    End Select
    
    Resume FunctionExit

End Function ' Cross

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Multiply vector by a scalar
'
' INPUT:        scalar - a number
'
' OUTPUT:       cVector object
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/30/2016   Created
'
'----------------------------------------------------------------------------------------
Public Function ScalarMultiply(ByVal scalar As Double) As cVector

    On Error GoTo ErrorHandler
    
    Dim result As cVector
    Set result = New cVector
    
    result.SetSize newSize:=pSize, bPreserve:=False
    
    Dim index As Long
    
    For index = 0 To pSize - 1
        result.Value(index) = scalar * pVector(index)
    Next index
    
    Set ScalarMultiply = result
    
FunctionExit:
    Exit Function
    
ErrorHandler:

    MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number, _
                   Buttons:=vbCritical
    
    Resume FunctionExit

End Function


'----------------------------------------------------------------------------------------
' CLASS PRIVATE METHODS
'----------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Kill objects in class
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/30/2016   Created
'
'----------------------------------------------------------------------------------------
Private Sub KillObject()
    pSize = 0
    Erase pVector
End Sub ' KillObject

