VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
' CLASS MODULE: cMatrix
'
' DESCRIPTION:  Matrix class for linear algebra calculations
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Option Explicit

'----------------------------------------------------------------------------------------
' CLASS DECLARATIONS
'----------------------------------------------------------------------------------------

' Error codes
Private Enum MatErr
    Add = vbObjectError + 513
    Subtract
    Multiply
    DeterminantCalc
    Inverse
    RowRange
    ColumnRange
    NotSquare
    Initialize
End Enum

' Error constants
Private Const MatrixErrorTitle As String = "Matrix Object Error"

Private pRows       As Long
Private pColumns    As Long
Private pMatrix()   As Double

'----------------------------------------------------------------------------------------
' CLASS EVENTS
'----------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    pRows = 0
    pColumns = 0
    ReDim pMatrix(0)
End Sub ' Class_Initialize

Private Sub Class_Terminate()
    KillObject
End Sub ' Class_Terminate

'----------------------------------------------------------------------------------------
' CLASS PROPERTIES
'----------------------------------------------------------------------------------------
Public Property Get Rows() As Long
    Rows = pRows
End Property ' rows(Get)

Public Property Get Columns() As Long
    Columns = pColumns
End Property ' columns(Get)

Public Property Get Value(ByVal row As Long, ByVal column As Long) As Double

    On Error GoTo ErrorHandler
    
    Const propName As String = "cMatrix.Value"
    
    ' check for out of range errors
    If row > UBound(pMatrix, 1) Then
    
        Err.Raise number:=MatErr.RowRange, _
                  source:=propName, _
                  description:="Row is out of range!"
                  
    ElseIf column > UBound(pMatrix, 2) Then
    
        Err.Raise number:=MatErr.ColumnRange, _
                  source:="cMatrix::Value", _
                  description:="Column is out of range!"
                  
    End If
    
    Value = pMatrix(row, column)
    
PropExit:
    Exit Property
    
ErrorHandler:

    Select Case Err.number
    
        Case MatErr.RowRange
            
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Error Source : " & Err.source, _
                   Buttons:=vbCritical, _
                   Title:=MatrixErrorTitle
            
        Case MatErr.ColumnRange
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Error Source : " & Err.source, _
                   Buttons:=vbCritical, _
                   Title:=MatrixErrorTitle
        
        Case Else
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number, _
                   Buttons:=vbCritical

    End Select
    
    Resume PropExit
    
End Property ' Value(Get)

Public Property Let Value(ByVal row As Long, ByVal column As Long, ByVal setValue As Double)

    On Error GoTo ErrorHandler
    
    Const propName As String = "cMatrix.Value"
    
    ' check for out of range errors
    If row > UBound(pMatrix, 1) Then
    
        Err.Raise number:=MatErr.RowRange, _
                  source:=propName, _
                  description:="Row is out of range!"
                  
    ElseIf column > UBound(pMatrix, 2) Then
    
        Err.Raise number:=MatErr.ColumnRange, _
                  source:=propName, _
                  description:="Column is out of range!"
                  
    End If
    
    pMatrix(row, column) = setValue
    
PropExit:
    Exit Property
    
ErrorHandler:

    Select Case Err.number
    
        Case MatErr.RowRange
            
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Source : " & Err.source, _
                   Buttons:=vbCritical, _
                   Title:=MatrixErrorTitle
            
        Case MatErr.ColumnRange
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Source : " & Err.source, _
                   Buttons:=vbCritical, _
                   Title:=MatrixErrorTitle
        
        Case Else
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number, _
                   Buttons:=vbCritical

    End Select
    
    Resume PropExit
 
End Property ' Value(Let)

'----------------------------------------------------------------------------------------
' CLASS PUBLIC METHODS
'----------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Sets the dimensions of the matrix
'
' INPUT:        rows - number of rows in matrix
'               column - number of columns in matrix
'               bPreserve - option to save current contents of matrix
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Public Sub SetDimensions(ByVal Rows As Long, ByVal Columns As Long, _
                Optional ByVal bPreserve As Boolean = False)
                
    ' if bPreserve = True, current contents of matrix will be maintained
    If bPreserve Then
        ReDim Preserve pMatrix(Rows - 1, Columns - 1)
    Else
        ReDim pMatrix(Rows - 1, Columns - 1)
    End If
    
    pRows = Rows
    pColumns = Columns

End Sub ' Dimensions

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Set matrix to the identity matrix
'
' INPUT:        newSize - size of matrix (n x n)
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    7/1/2016   Created
'
'----------------------------------------------------------------------------------------
Public Sub IdentityMatrix(ByVal newSize As Long)
    
    Me.SetDimensions Rows:=newSize, Columns:=newSize, bPreserve:=False
    
    Dim diag As Long
    For diag = 0 To pRows - 1
        pMatrix(diag, diag) = 1
    Next diag
    
End Sub ' IdentityMatrix

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Populates the entries of the matrix from an array of values in the
'               spreadsheet of unkown size
'
' INPUTS:       startRange - top left cell in matrix
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Public Sub DynamicMatrixFromSheet(ByVal startRange As Range)

    On Error GoTo ErrorHandler
    
    Const subName As String = "cMatrix.DynamicMatrixFromSheet"
    
    Dim shRows As Long
    Dim shColumns As Long
    
    ' find size of matrix by counting cells until empty cell is found
    With startRange.CurrentRegion
        shRows = .Rows.Count
        shColumns = .Columns.Count
    End With
    
    ' resize range
    Set startRange = startRange.Resize(shRows, shColumns)
    
    ' resize matrix array
    Me.SetDimensions Rows:=shRows, Columns:=shColumns
    
    ' populate matrix
    Dim row As Long
    Dim column As Long
    
    For row = 0 To pRows - 1
        For column = 0 To pColumns - 1
            ' note that excel ranges are base 1 indexed
            Me.Value(row, column) = startRange(row + 1, column + 1)
        Next column
    Next row
    
SubExit:
    Exit Sub
    
ErrorHandler:

    MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                   "Error Number : " & Err.number, _
           Buttons:=vbCritical
    
    Resume SubExit
    
End Sub ' DynamicMatrixFromSheet

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Populates the entries of the matrix from an array of values in the
'               spreadsheet of a known size
'
' INPUTS:       matrixRange - array of cells in spreadsheet encompasing matrix
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Public Sub StaticMatrixFromSheet(ByVal matrixRange As Range)

    Dim shRows As Long
    Dim shColumns As Long
    
    With matrixRange
        shRows = .Rows.Count
        shColumns = .Columns.Count
    End With
    
    ' resize matrix array
    Me.SetDimensions Rows:=shRows, Columns:=shColumns
    
    ' populate matrix
    Dim row As Long
    Dim column As Long
    
    For row = 0 To pRows - 1
        For column = 0 To pColumns - 1
            ' note that excel ranges are base 1 indexed
            Me.Value(row, column) = matrixRange(row + 1, column + 1)
        Next column
    Next row

SubExit:
    Exit Sub
    
ErrorHandler:

    MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                   "Error Number : " & Err.number, _
           Buttons:=vbCritical
    
    Resume SubExit
    
End Sub ' StaticMatrixFromSheet

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Prints matrix to Immediate Window
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Public Sub PrintToConsole()

    On Error GoTo ErrorHandler

    Dim rowString As String
    rowString = vbNullString
    
    Dim row As Long
    Dim column As Long
    For row = 0 To pRows - 1
        rowString = CStr(pMatrix(row, 0))
        For column = 1 To pColumns - 1
            rowString = rowString & vbTab & CStr(pMatrix(row, column))
        Next column
        Debug.Print rowString
    Next row
    
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
' DESCRIPTION:  Outputs matrix to spreadsheet
'
' INPUT:        startCell - location of first cell to print matrix to
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Public Sub OutputToSpreadsheet(ByVal startCell As Range)

    On Error GoTo ErrorHandler

    startCell.Resize(pRows, pColumns).Value2 = pMatrix
    
SubExit:
    Exit Sub
    
ErrorHandler:

    MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                   "Error Number : " & Err.number, _
           Buttons:=vbCritical
    
    Resume SubExit
    
End Sub ' OutputToSpreadsheet

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Outputs matrix to spreadsheet
'
' INPUT:        startCell - location of first cell to print matrix to
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Public Sub OutPutToFile(ByVal fileName As String)

    On Error GoTo ErrorHandler
    
    Dim folderPath As String
    folderPath = Application.ThisWorkbook.Path
    
    Dim writeFile As String
    writeFile = folderPath & "\" & fileName & ".txt"
    
    Dim rowString As String
    
    Open writeFile For Output As #1
    
    rowString = vbNullString
    
    Dim row As Long
    Dim column As Long
    
    For row = 0 To pRows - 1
    
        rowString = CStr(pMatrix(row, 0))
        
        For column = 1 To pColumns - 1
        
            rowString = rowString & vbTab & CStr(pMatrix(row, column))
            
        Next column
        
        Print #1, rowString
        
    Next row
    
SubExit:

    Close #1
    
    Exit Sub
    
ErrorHandler:

    MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                   "Error Number : " & Err.number, _
           Buttons:=vbCritical
    
    Resume SubExit

End Sub ' OutPutToFile

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Swaps two rows in the matrix
'
' INPUTS:       row1 - first row to swap
'               row2 - second row to swap
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Public Sub SwapRows(ByVal row1 As Long, ByVal row2 As Long)

    On Error GoTo ErrorHandler
    
    Const subName As String = "cMatrix.SwapRows"
    
    ' check for out of range errors
    If row1 > UBound(pMatrix, 1) Or row2 > UBound(pMatrix, 1) Then
    
        Err.Raise number:=MatErr.RowRange, _
                  source:=subName, _
                  description:="Row is out of range!"
                  
    End If

    Dim temp As Double
    Dim column As Long
    
    For column = 0 To pColumns - 1
        temp = Me.Value(row2, column)
        Me.Value(row2, column) = Me.Value(row1, column)
        Me.Value(row1, column) = temp
    Next column
    
SubExit:
    Exit Sub
    
ErrorHandler:

    Select Case Err.number
    
        Case MatErr.RowRange
            
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Source : " & Err.source, _
                   Buttons:=vbCritical, _
                   Title:=MatrixErrorTitle
        
        Case Else
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number, _
                   Buttons:=vbCritical

    End Select
    
    Resume SubExit
    
End Sub ' SwapRows

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Creates a copy of the current matrix object
'
' OUTPUTS:      cMatrix object
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Public Function Copy() As cMatrix

    On Error GoTo ErrorHandler

    Dim result As cMatrix
    Set result = New cMatrix
    
    result.SetDimensions Rows:=pRows, Columns:=pColumns
    
    Dim row As Long
    Dim column As Long
    For row = 0 To pRows - 1
        For column = 0 To pColumns - 1
            result.Value(row, column) = Me.Value(row, column)
        Next column
    Next row
    
    Set Copy = result
    
FunctionExit:
    
    Exit Function
    
ErrorHandler:
    
    MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Source : " & Err.source, _
                   Buttons:=vbCritical
                   
    Set result = Nothing
    
    Resume FunctionExit
    
End Function ' Copy

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Multiply all elements of matrix by a scalar
'
' INPUTS:       scalar - a scalar number
'
' OUTPUTS:      cMatrix object
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Public Function ScalarMultiply(ByVal scalar As Double) As cMatrix

    On Error GoTo ErrorHandler
    
    Dim result As cMatrix
    Set result = New cMatrix
    
    result.SetDimensions Rows:=pRows, Columns:=pColumns
    
    Dim row As Long
    Dim column As Long
    For row = 0 To pRows - 1
        For column = 0 To pColumns - 1
            result.Value(row, column) = scalar * Me.Value(row, column)
        Next column
    Next row
    
    Set ScalarMultiply = result
    
FunctionExit:
    
    Exit Function
    
ErrorHandler:
    
    MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Source : " & Err.source, _
                   Buttons:=vbCritical
    
    Set result = Nothing
    
    Resume FunctionExit
    
End Function ' ScalarMultiply

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Multiply matrix by another matrix
'
' INPUTS:       otherMatrix - cMatrix object of compatible size for multiplication
'
' OUTPUTS:      cMatrix object
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Public Function MatrixMultiply(ByVal otherMatrix As cMatrix) As cMatrix

    On Error GoTo ErrorHandler
    
    Const functionName As String = "cMatrix.MatrixMultiply"
    
    ' check that matrices can me multiplied
    If pColumns <> otherMatrix.Rows Then
        
        Err.Raise number:=MatErr.Multiply, _
                  source:=functionName, _
                  description:="Matrices not compitable for multiplication!"

    End If
    
    Dim result As cMatrix
    Set result = New cMatrix
    
    result.SetDimensions Rows:=pRows, Columns:=otherMatrix.Columns
    
    Dim rowA As Long
    Dim columnB As Long
    Dim columnA As Long
    
    For rowA = 0 To pRows - 1
        For columnB = 0 To otherMatrix.Columns - 1
            For columnA = 0 To pColumns - 1
            
                result.Value(rowA, columnB) = result.Value(rowA, columnB) + _
                    Me.Value(rowA, columnA) * otherMatrix.Value(columnA, columnB)
                    
            Next columnA
        Next columnB
    Next rowA
    
    Set MatrixMultiply = result
    
FunctionExit:

    Exit Function
    
ErrorHandler:
    
    Select Case Err.number
    
        Case MatErr.Multiply
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & vbNewLine & _
                           "Source : " & Err.source, _
                   Buttons:=vbCritical, _
                   Title:=MatrixErrorTitle
        
        Case Else
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Source : " & Err.source, _
                   Buttons:=vbCritical
    
    End Select
    
    Set result = Nothing

    Resume FunctionExit
    
End Function ' MatrixMultiply

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Add matrix to another matrix
'
' INPUTS:       otherMatrix - cMatrix object of compatible size for multiplication
'
' OUTPUTS:      cMatrix object
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Public Function MatrixAdd(ByVal otherMatrix As cMatrix) As cMatrix
    
    On Error GoTo ErrorHandler
    
    Const functionName As String = "cMatrix.MatrixAdd"
    
    If Not pRows = otherMatrix.Rows Or Not pColumns = otherMatrix.Columns Then
        
        Err.Raise number:=MatErr.Add, _
                  source:=functionName, _
                  description:="Matrices not compitable for addition!"
        
    End If
    
    Dim result As cMatrix
    Set result = New cMatrix
    
    result.SetDimensions Rows:=pRows, Columns:=pColumns
    
    Dim row As Long
    Dim column As Long
    
    For row = 0 To pRows - 1
        For column = 0 To pColumns - 1
            result.Value(row, column) = Me.Value(row, column) + _
                otherMatrix.Value(row, column)
        Next column
    Next row
    
    Set MatrixAdd = result

FunctionExit:

    Exit Function
    
ErrorHandler:
    
    Select Case Err.number
    
        Case MatErr.Add
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & vbNewLine & _
                           "Source : " & Err.source, _
                   Buttons:=vbCritical, _
                   Title:=MatrixErrorTitle
        
        Case Else
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Source : " & Err.source, _
                   Buttons:=vbCritical
    
    End Select
    
    Set result = Nothing

    Resume FunctionExit

End Function ' MatrixAdd

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Subtract matrix by another matrix
'
' INPUTS:       otherMatrix - cMatrix object of compatible size for multiplication
'
' OUTPUTS:      cMatrix object
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Public Function MatrixSubtract(ByVal otherMatrix As cMatrix) As cMatrix
    
    On Error GoTo ErrorHandler
    
    Const functionName As String = "cMatrix.MatrixSubtract"
    
    If Not pRows = otherMatrix.Rows Or Not pColumns = otherMatrix.Columns Then
        
        Err.Raise number:=MatErr.Subtract, _
                  source:=functionName, _
                  description:="Matrices not compitable for subtraction!"
        
    End If
    
    Dim result As cMatrix
    Set result = New cMatrix
    
    result.SetDimensions Rows:=pRows, Columns:=pColumns
    
    Dim row As Long
    Dim column As Long
    
    For row = 0 To pRows - 1
        For column = 0 To pColumns - 1
            result.Value(row, column) = Me.Value(row, column) - _
                otherMatrix.Value(row, column)
        Next column
    Next row
    
    Set MatrixSubtract = result

FunctionExit:

    Exit Function
    
ErrorHandler:
    
    Select Case Err.number
    
        Case MatErr.Subtract
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & vbNewLine & _
                           "Source : " & Err.source, _
                   Buttons:=vbCritical, _
                   Title:=MatrixErrorTitle
        
        Case Else
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Source : " & Err.source, _
                   Buttons:=vbCritical
    
    End Select
    
    Set result = Nothing

    Resume FunctionExit

End Function ' MatrixSubtract

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Transpose matrix
'
' OUTPUTS:      cMatrix object
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Public Function Transpose() As cMatrix

    On Error GoTo ErrorHandler
    
    Dim result As cMatrix
    Set result = New cMatrix
    
    result.SetDimensions Rows:=pColumns, Columns:=pRows
    
    Dim row As Long
    Dim column As Long
    
    For row = 0 To pColumns - 1
        For column = 0 To pRows - 1
        
            result.Value(row, column) = Me.Value(column, row)
        
        Next column
    Next row
    
    Set Transpose = result
    
FunctionExit:
    
    Exit Function
    
ErrorHandler:
    
    MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Source : " & Err.source, _
                   Buttons:=vbCritical
                   
    Set result = Nothing
    
    Resume FunctionExit

End Function ' Transpose

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Check if matrix is square
'
' OUTPUT:       True or False
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Public Function IsSquareMatrix() As Boolean

    If pRows <> pColumns Then
        IsSquareMatrix = False
    ElseIf pRows = pColumns Then
        IsSquareMatrix = True
    End If

End Function ' IsSquareMatrix

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Calculate determinant of matrix recursively
'
' OUTPUT:       Determinant
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Public Function Determinant() As Double

    On Error GoTo ErrorHandler
    
    Const functionName As String = "cMatrix.Determinant"
    
    If Not Me.IsSquareMatrix Then
        
        Err.Raise number:=MatErr.DeterminantCalc, _
                  source:=functionName, _
                  description:="Cannot calculate determinant of a non-square matrix"
        
    End If
    
    Dim result As Double
    result = 0#
    
    Dim i_outer As Long
    Dim i As Long
    Dim j As Long
    
    Dim subMatrix As cMatrix
    
    If pRows = 1 Then
        result = Me.Value(0, 0)
    Else
    
        For i_outer = 0 To pRows - 1
        
            Set subMatrix = New cMatrix
            subMatrix.SetDimensions Rows:=(pRows - 1), Columns:=(pColumns - 1)
            
            For i = 0 To pRows - 2
                For j = 0 To i_outer - 1
                    subMatrix.Value(i, j) = Me.Value(i + 1, j)
                Next j
                
                For j = i_outer To pRows - 2
                    subMatrix.Value(i, j) = Me.Value(i + 1, j + 1)
                Next j
            Next i
            
            result = result + (-1#) ^ i_outer _
                * Me.Value(0, i_outer) * subMatrix.Determinant
            
        Next i_outer
    
    End If
    
    Determinant = result
    
FunctionExit:

    Exit Function
    
ErrorHandler:

    Select Case Err.number
    
        Case MatErr.DeterminantCalc
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & vbNewLine & _
                           "Source : " & Err.source, _
                   Buttons:=vbCritical, _
                   Title:=MatrixErrorTitle
                   
        Case Else
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & vbNewLine & _
                           "Source : " & Err.source, _
                   Buttons:=vbCritical
    
    End Select
    
    Determinant = 0
    
    Resume FunctionExit

End Function ' Determinant

'----------------------------------------------------------------------------------------
' CLASS PRIVATE METHODS
'----------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Kill objects in class
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Private Sub KillObject()
    pRows = 0
    pColumns = 0
    Erase pMatrix
End Sub ' KillObject
