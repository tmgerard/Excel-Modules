VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cRebar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' CLASS MODULE: cRebar
'
' DESCRIPTION:  Class for creating reinforcing objects
'
' TODO:         Error handling
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/15/2016   Created
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' CLASS DECLARATIONS
'------------------------------------------------------------------------------

Public Enum RebarSize ' Enumeration for bar selection
    three = 3
    four
    five
    six
    seven
    eight
    nine
    ten
    eleven
    fourteen = 14
    eighteen = 18
End Enum

Private pBarSize        As RebarSize
Private pWeight         As Double       ' pounds per foot
Private pArea           As Double       ' inches
Private pDiameter       As Double       ' square inches
Private pPerimeter      As Double       ' inches
Private pEpoxyCoated    As Boolean

'------------------------------------------------------------------------------
' CLASS EVENTS
'------------------------------------------------------------------------------

Private Sub Class_Initialize()
    pBarSize = RebarSize.four
    pEpoxyCoated = False
End Sub

'------------------------------------------------------------------------------
' CLASS PROPERTIES
'------------------------------------------------------------------------------
Public Property Get BarSize() As RebarSize
    BarSize = pBarSize
End Property

Public Property Let BarSize(ByVal size As RebarSize)
    pBarSize = size
    UpdateProperties
End Property

Public Property Get Weight() As Double
    Weight = pWeight
End Property

Public Property Get Area() As Double
    Area = pArea
End Property

Public Property Get Diameter() As Double
    Diameter = pDiameter
End Property

Public Property Get Perimeter() As Double
    Perimeter = pPerimeter
End Property

Public Property Get IsEpoxyCoated() As Boolean
    IsEpoxyCoated = pEpoxyCoated
End Property

Public Property Let IsEpoxyCoated(ByVal isCoated As Boolean)
    pEpoxyCoated = isCoated
End Property

'------------------------------------------------------------------------------
' CLASS METHODS
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DESCRIPTION:  Called in methods that change pBarSize to update the rebar
'               properties stored in the object based on what is stored
'               in pBarSize.
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/15/2016   Created
'
'------------------------------------------------------------------------------
Private Sub UpdateProperties()

    Select Case pBarSize
    
        Case RebarSize.three
            pWeight = 0.376
            pDiameter = 0.375
            pArea = 0.11
            pPerimeter = 1.178

        Case RebarSize.four
            pWeight = 0.668
            pDiameter = 0.5
            pArea = 0.2
            pPerimeter = 1.571

        Case RebarSize.five
            pWeight = 1.043
            pDiameter = 0.625
            pArea = 0.31
            pPerimeter = 1.963

        Case RebarSize.six
            pWeight = 1.502
            pDiameter = 0.75
            pArea = 0.44
            pPerimeter = 2.356

        Case RebarSize.seven
            pWeight = 2.044
            pDiameter = 0.875
            pArea = 0.6
            pPerimeter = 2.749

        Case RebarSize.eight
            pWeight = 2.67
            pDiameter = 1#
            pArea = 0.79
            pPerimeter = 3.142

        Case RebarSize.nine
            pWeight = 3.4
            pDiameter = 1.128
            pArea = 1#
            pPerimeter = 3.544

        Case RebarSize.ten
            pWeight = 4.303
            pDiameter = 1.27
            pArea = 1.27
            pPerimeter = 3.99

        Case RebarSize.eleven
            pWeight = 5.313
            pDiameter = 1.41
            pArea = 1.56
            pPerimeter = 4.43

        Case RebarSize.fourteen
            pWeight = 7.65
            pDiameter = 1.693
            pArea = 2.25
            pPerimeter = 5.32

        Case RebarSize.eighteen
            pWeight = 13.6
            pDiameter = 2.257
            pArea = 4#
            pPerimeter = 7.09
    
    End Select
    
End Sub
