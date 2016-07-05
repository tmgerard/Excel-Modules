VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLinearSolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
' CLASS MODULE: cLinearSolver
'
' DESCRIPTION:  Class for solving systems of linear equations [Ax = b]
'
' REQUIRES:     cMatrix and cVector Modules
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
Private Enum SolverErr
    Incompatible = vbObjectError + 528
    MatrixNotSquare
    SingularMatrix
End Enum

' Error constants
Private Const LSErrorTitle = "Linear Solver Object Error"

Private pMatrixA    As cMatrix
Private pVectorb    As cVector
Private pSolution   As cVector

'----------------------------------------------------------------------------------------
' CLASS EVENTS
'----------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set pMatrixA = New cMatrix
    Set pVectorb = New cVector
    Set pSolution = New cVector
End Sub ' Class_Initialize

Private Sub Class_Terminate()
    KillObjects
End Sub ' Class_Terminate

'----------------------------------------------------------------------------------------
' CLASS PROPERTIES
'----------------------------------------------------------------------------------------



'----------------------------------------------------------------------------------------
' CLASS PUBLIC METHODS
'----------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------
' CLASS PUBLIC METHODS
'----------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Sets up system of linear equations to solve
'
' INPUT:        matrix - the coefficient matrix A
'
'               vector - the right hand size vector b
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/30/2016   Created
'
'----------------------------------------------------------------------------------------
Public Sub CreateLinearSystem(ByVal matrix As cMatrix, ByVal vector As cVector)

    On Error GoTo ErrorHandler
    
    Const subName As String = "cLinearSolver.CreateLinearSystem"
    
    If Not matrix.IsSquareMatrix Then
    
        Err.Raise number:=SolverErr.MatrixNotSquare, _
                  source:=subName, _
                  description:="Matrix is not square!"
    
    ElseIf matrix.Rows <> vector.Size Then
    
        Err.Raise number:=SolverErr.Incompatible, _
                  source:=subName, _
                  description:="Matrix and vector sizes incompatible!"
    
    End If
    
    Set pMatrixA = matrix
    Set pVectorb = vector
    pSolution.SetSize newSize:=vector.Size
    
SubExit:
    
    Exit Sub
    
ErrorHandler:

    Select Case Err.number
    
        Case SolverErr.MatrixNotSquare
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Error Source : " & Err.source, _
                   Buttons:=vbCritical, _
                   Title:=LSErrorTitle
        
        Case SolverErr.Incompatible
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Error Source : " & Err.source, _
                   Buttons:=vbCritical, _
                   Title:=LSErrorTitle
        
        Case Else
        
            MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & _
                           "Error Source : " & Err.source, _
                   Buttons:=vbCritical
    
    End Select
    
    
    Resume SubExit

End Sub ' CreateLinearSystem

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Calculates lower and upper triangular matrices for LU Decomposition.
'               No pivoting means this is error prone.
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/30/2016   Created
'
'----------------------------------------------------------------------------------------
Public Sub LU_Decomposition()

    On Error GoTo ErrorHandler
    
    Dim A As cMatrix
    Set A = pMatrixA.Copy
    
    Dim n As Long
    n = A.Rows
    
    Dim U As cMatrix
    Set U = New cMatrix
    U.SetDimensions Rows:=n, Columns:=n, bPreserve:=False
    
    Dim L As cMatrix
    Set L = New cMatrix
    L.IdentityMatrix newSize:=n
    
    Dim k As Long
    Dim i As Long
    Dim j As Long
    
    For k = 0 To n - 1
    
        U.Value(k, k) = A.Value(k, k)
        
        For i = k + 1 To n - 1
            L.Value(i, k) = A.Value(i, k) / U.Value(k, k)
            U.Value(k, i) = A.Value(k, i)
        Next i
        
        For i = k + 1 To n - 1
        
            For j = k + 1 To n - 1
            
                A.Value(i, j) = A.Value(i, j) - L.Value(i, k) * U.Value(k, j)
            
            Next j
        
        Next i
        
    Next k
    
    L.PrintToConsole
    U.PrintToConsole
    
    
    
SubExit:
    
    Set A = Nothing
    Set U = Nothing
    Set L = Nothing
    Exit Sub
    
ErrorHandler:

    MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                           "Error Number : " & Err.number & vbNewLine & _
                           "Error Source : " & Err.source, _
                   Buttons:=vbCritical
    
    Resume SubExit

End Sub ' LU_Decomposition

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Calculates lower triangular, upper triangular, and permutation matrices
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/30/2016   Created
'
'----------------------------------------------------------------------------------------
Public Sub LUP_Decomposition()

    On Error GoTo ErrorHandler
    
    Const subName As String = "cLinearSolver.LUP_Decomposition"

    Dim A As cMatrix
    Set A = pMatrixA.Copy
    
    Dim n As Long
    n = A.Rows
    
    Dim pArray As cVector
    Set pArray = New cVector
    pArray.SetSize newSize:=n, bPreserve:=False
        
    Dim i As Long
    Dim k As Long
    Dim j As Long
    Dim kPrime As Long
    Dim p As Double
    Dim temp As Double
    
    ' initialize permutation array
    For i = 0 To n - 1
        pArray.Value(i) = i
    Next i
    
    For k = 0 To n - 1
    
        p = 0
        
        For i = k To n - 1
        
            If Abs(A.Value(i, k)) > p Then
            
                p = Abs(A.Value(i, k))
                kPrime = i
            
            End If
        
        Next i
        
        If p = 0 Then
        
            Err.Raise number:=SolverErr.SingularMatrix, _
                  source:=subName, _
                  description:="Singular matrix"
        
        End If

        ' pivot matrix and permutation array
        pArray.SwapValues index1:=k, index2:=kPrime
        A.SwapRows row1:=k, row2:=kPrime
        
        For i = k + 1 To n - 1
        
            A.Value(i, k) = A.Value(i, k) / A.Value(k, k)
            
            For j = k + 1 To n - 1
            
                A.Value(i, j) = A.Value(i, j) - A.Value(i, k) * A.Value(k, j)
            
            Next j
        
        Next i
    
    Next k
    
    'L and U stored in original copy of A
    A.PrintToConsole
    pArray.PrintToConsole
    
    ' sovle system
    Dim y As cVector
    Set y = New cVector
    y.SetSize newSize:=n, bPreserve:=False
    
    Dim sum As Double
    
    ' forward substitusion
    For i = 0 To n - 1
        sum = 0
        For j = 0 To i
            sum = sum + A.Value(i, j) * y.Value(j)
        Next j
        
        y.Value(i) = pVectorb.Value(pArray.Value(i)) - sum
        
    Next i
    
    y.PrintToConsole
    
    ' back substitution
    For i = n - 1 To 0 Step -1
        sum = 0
        For j = i To n - 1
            sum = sum + A.Value(i, j) * pSolution.Value(j)
        Next j
        
        pSolution.Value(i) = (y.Value(i) - sum) / A.Value(i, i)
        
    Next i
    
    pSolution.PrintToConsole
    
SubExit:

    Set A = Nothing
    Set pArray = Nothing
    
    Exit Sub
    
ErrorHandler:

    Select Case Err.number
    
    Case SolverErr.SingularMatrix
    
        MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                               "Error Number : " & Err.number & vbNewLine & _
                               "Error Source : " & Err.source, _
                       Buttons:=vbCritical, _
                       Title:=LSErrorTitle
    
    Case Else

        MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                               "Error Number : " & Err.number & vbNewLine & _
                               "Error Source : " & Err.source, _
                       Buttons:=vbCritical
    
    End Select
    
    Resume SubExit

End Sub ' LUP_Decomposition

'----------------------------------------------------------------------------------------
' DESCRIPTION:  Outputs solution vector to spreadsheet
'
' INPUT:        startCell - location of first cell to print matrix to
'
' AUTHOR        DATE        NOTES
'----------------------------------------------------------------------------------------
' Tom Gerard    6/28/2016   Created
'
'----------------------------------------------------------------------------------------
Public Sub OutputSolutionToSpreadsheet(ByVal startCell As Range, _
                                       Optional ByVal vertical As Boolean = True)

    On Error GoTo ErrorHandler

    pSolution.OutputToSpreadsheet startCell, vertical
    
SubExit:
    Exit Sub
    
ErrorHandler:

    MsgBox Prompt:="Error Message : " & Err.description & vbNewLine & _
                   "Error Number : " & Err.number, _
           Buttons:=vbCritical
    
    Resume SubExit
    
End Sub ' OutputToSpreadsheet

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
Private Sub KillObjects()
    Set pMatrixA = Nothing
    Set pVectorb = Nothing
    Set pSolution = Nothing
End Sub ' KillObjects
