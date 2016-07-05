VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTensionRebarDevelop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' CLASS MODULE: cTensionRebarDevelop
'
' DESCRIPTION:  Class containing calculations for rebar development length for
'               bars in tension per Section 5.11.2.1 [pg. 5-160] per AASHTO
'               LRFD Bridge Design Specification, 6th Edition, with 2013
'               Interim Revisions
'
' REQUIREMENTS: cRebar class module
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

Private pBar                                As cRebar
Private pSteelYield                         As Double   ' ksi
Private pConcreteCompStrength               As Double   ' ksi
Private pLightWeightFct                     As Double   ' ksi
Private pAsProvided                         As Double   ' sq. in.
Private pAsRequired                         As Double   ' sq. in.

' Modificatin factor flags
Private p12_InchesBelowBar                  As Boolean
Private pFctLWConcrete                      As Boolean
Private pLWConcrete                         As Boolean
Private pSandLWConcrete                     As Boolean
Private pEpoxyCoatedCoverAndSpacing         As Boolean
Private pEpoxyCoatedOther                   As Boolean
Private pCoverAndSpacing                    As Boolean
Private pAreaSteelUsed                      As Boolean
Private pEnclosedInSpiral                   As Boolean

' Calculated modification factors
Private pFctFactor                          As Double
Private pAreaSteelFactor                    As Double
Private pTotalIncreaseFactor                As Double
Private pTotalDecreaseFactor                As Double

' Constant modification factors that increase development length
' from Section 5.11.2.1.2 of the AASHTO LRFD Bridge Design Specifications,
' 6th Edition, with 2013 Interim Revisions.
Private Const BAR12INCHESFRESHBELOW         As Double = 1.4
Private Const LIGHTWEIGHTCONCRETENOFCT      As Double = 1.3
Private Const SANDLIGHTWEIGHTCONCRETENOFCT  As Double = 1.2
Private Const EPOXYCOATEDCOVERORSPACING     As Double = 1.5
Private Const EPOXYCOATEDCATCHALL           As Double = 1.2

' Constant modification factors that decrease development length
' from Section 5.11.2.1.3 of the AASHTO LRFD Bridge Design Specifications,
' 6th Edition, with 2013 Interim Revisions.
Private Const COVERANDLATERALSPACING        As Double = 0.8
Private Const ENCLOSEDINSPIRALFACTOR        As Double = 0.75

'------------------------------------------------------------------------------
' CLASS EVENTS
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set pBar = New cRebar
    pSteelYield = 60
    pConcreteCompStrength = 3.5
    pLightWeightFct = 0#
    pAsProvided = 0#
    pAsRequired = 0#
    
    ' initialize all modification factor flags to false
    p12_InchesBelowBar = False
    pFctLWConcrete = False
    pLWConcrete = False
    pSandLWConcrete = False
    pEpoxyCoatedCoverAndSpacing = False
    pEpoxyCoatedOther = False
    pCoverAndSpacing = False
    pAreaSteelUsed = False
    pEnclosedInSpiral = False
    
    ' initialize modification factors
    pFctFactor = 1#
    pAreaSteelFactor = 1#
    pTotalIncreaseFactor = 1#
    pTotalDecreaseFactor = 1#

End Sub

Private Sub Class_Terminate()
    KillObjects
End Sub

'------------------------------------------------------------------------------
' CLASS PROPERTIES
'------------------------------------------------------------------------------
Public Property Get BarSize() As RebarSize
    BarSize = pBar.BarSize
End Property

Public Property Let BarSize(ByVal size As RebarSize)
    pBar.BarSize = size
End Property

Public Property Get SteelYield() As Double
    SteelYield = pSteelYield
End Property

Public Property Let SteelYield(ByVal yield As Double)
    pSteelYield = yield
End Property

Public Property Get ConcreteCompressiveStrength() As Double
    ConcreteCompressiveStrength = pConcreteCompStrength
End Property

Public Property Let ConcreteCompressiveStrength(ByVal compStrength As Double)
    pConcreteCompStrength = compStrength
End Property

Public Property Get BarWith12InchConcreteBelow() As Boolean
    BarWith12InchConcreteBelow = p12_InchesBelowBar
End Property

Public Property Let BarWith12InchConcreteBelow(ByVal flag As Boolean)
    p12_InchesBelowBar = flag
End Property

Public Property Get StdLightWeightConcrete() As Boolean
    StdLightWeightConcrete = pLWConcrete
End Property

Public Property Let StdLightWeightConcrete(ByVal flag As Boolean)
    ' Disable the other lightweight concrete factor
    If flag = True Then
        pFctLWConcrete = False
        pSandLWConcrete = False
    End If
    
    pLWConcrete = flag
End Property

Public Property Get SandLightWeightConcrete() As Boolean
    SandLightWeightConcrete = pSandLWConcrete
End Property

Public Property Let SandLightWeightConcrete(ByVal flag As Boolean)
    'Disable the other lightweight concrete factor
    If flag = True Then
        pFctLWConcrete = False
        pLWConcrete = False
    End If
    
    pSandLWConcrete = flag
End Property

Public Property Get FctLightWeightConcrete() As Boolean
    FctLightWeightConcrete = pFctLWConcrete
End Property

Public Property Let FctLightWeightConcrete(ByVal flag As Boolean)
    
    If flag = True Then
        pLWConcrete = False
        pSandLWConcrete = False
    End If

    pFctLWConcrete = flag
End Property

Public Property Get Fct() As Double
    Fct = pLightWeightFct
End Property

Public Property Let Fct(ByVal FctValue As Double)

    If FctValue > 0 Then
        pLightWeightFct = FctValue
    Else
        pLightWeightFct = 0#
    End If
    
    CalculateL_W_GivenFct
    
End Property

Public Property Get EpoxyCoatedCoverAndSpacing() As Boolean
    EpoxyCoatedCoverAndSpacing = pEpoxyCoatedCoverAndSpacing
End Property

Public Property Let EpoxyCoatedCoverAndSpacing(ByVal flag As Boolean)
    ' disable the other epoxy coating flag
    If flag = True Then
        pEpoxyCoatedOther = False
    End If
    
    pEpoxyCoatedCoverAndSpacing = flag ' set for cRebar object
End Property

Public Property Get EpoxyCoatedOther() As Boolean
    EpoxyCoatedOther = pEpoxyCoatedOther
End Property

Public Property Let EpoxyCoatedOther(ByVal flag As Boolean)
    ' disable the other epoxy coating flag
    If flag = True Then
        pEpoxyCoatedCoverAndSpacing = False
    End If
    
    pEpoxyCoatedOther = flag
End Property

Public Property Get CoverAndSpacing() As Boolean
    CoverAndSpacing = pCoverAndSpacing
End Property

Public Property Let CoverAndSpacing(ByVal flag As Boolean)
    pCoverAndSpacing = flag
End Property

Public Property Get AreaSteelUsed() As Boolean
    AreaSteelUsed = pAreaSteelUsed
End Property

Public Property Let AreaSteelUsed(ByVal flag As Boolean)
    pAreaSteelUsed = flag
End Property

Public Property Get EnclosedInSpiral() As Boolean
    EnclosedInSpiral = pEnclosedInSpiral
End Property

Public Property Let EnclosedInSpiral(ByVal flag As Boolean)
    pEnclosedInSpiral = flag
End Property

Public Property Get IncreasingFactor() As Double
    CalculateIncreaseModificationFactor
    IncreasingFactor = pTotalIncreaseFactor
End Property

Public Property Get DecreasingFactor() As Double
    CalculateDecreaseModificationFactor
    DecreasingFactor = pTotalDecreaseFactor
End Property

'------------------------------------------------------------------------------
' CLASS METHODS
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DESCRIPTION:  Section 5.11.2.1.1 [pg. 5-160] in the AASHTO LRFD Bridge Design
'               Specifications, 6th Edition, with 2013 Interim Revisions.
'
' OUTPUT:       Basic development length
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/15/2016   Created
'
'------------------------------------------------------------------------------
Public Function BasicDevelopmentLength() As Double

    If pBar.BarSize = RebarSize.eighteen Then
        BasicDevelopmentLength = BDL18
    ElseIf pBar.BarSize = RebarSize.fourteen Then
        BasicDevelopmentLength = BDL14
    Else ' For No. 11 bar and smaller
        BasicDevelopmentLength = BDL11_AndSmaller
    End If
    
End Function

'------------------------------------------------------------------------------
' DESCRIPTION:  Helper function for BasicDevelopmentLength
'
' OUTPUT:       Basic development length for No. 18 bar
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/15/2016   Created
'
'------------------------------------------------------------------------------
Private Function BDL18() As Double
    BDL18 = (3.5 * pSteelYield) / Math.Sqr(pConcreteCompStrength)
End Function

'------------------------------------------------------------------------------
' DESCRIPTION:  Helper function for BasicDevelopmentLength
'
' OUTPUT:       Basic development length for No. 14 bar
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/15/2016   Created
'
'------------------------------------------------------------------------------
Private Function BDL14() As Double
    BDL14 = (2.7 * pSteelYield) / Math.Sqr(pConcreteCompStrength)
End Function

'------------------------------------------------------------------------------
' DESCRIPTION:  Helper function for BasicDevelopmentLength
'
' OUTPUT:       Basic development length for No. 11 and smaller bars
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/15/2016   Created
'
'------------------------------------------------------------------------------
Private Function BDL11_AndSmaller() As Double
    
    Dim basicDevLength As Double
    Dim limit As Double
    
    basicDevLength = (1.25 * pBar.Area * pSteelYield) _
                        / Math.Sqr(pConcreteCompStrength)
                        
     ' lower limit to development length
    limit = 0.4 * pBar.Diameter * pSteelYield

    If basicDevLength < limit Then
        BDL11_AndSmaller = limit
    Else
        BDL11_AndSmaller = basicDevLength
    End If
    
End Function

'------------------------------------------------------------------------------
' DESCRIPTION:  Modified developmentlength
'
' OUTPUT:       Basic development length multiplied by modification factors
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/16/2016   Created
'
'------------------------------------------------------------------------------
Public Function ModifiedDevelopment() As Double
    ' make sure modifications factors calculated
    CalculateIncreaseModificationFactor
    CalculateDecreaseModificationFactor
    
    ModifiedDevelopment = Me.BasicDevelopmentLength * pTotalIncreaseFactor _
        * pTotalDecreaseFactor
End Function

'------------------------------------------------------------------------------
' DESCRIPTION:  Set area of steel for reduction factor
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/16/2016   Created
'
'------------------------------------------------------------------------------
Public Sub SetAreaOfSteelProvidedAndRequired(ByVal AsProvided As Double, _
                                             ByVal AsRequired As Double)
    If AsProvided <= 0 Or AsRequired <= 0 Then
        pAsProvided = 1#
        pAsRequired = 1#
    ElseIf AsProvided < AsRequired Then
        MsgBox "Area steel provided must be greater than Area of steel required!", _
            vbExclamation
        pAsProvided = 1#
        pAsRequired = 1#
    Else
        pAsProvided = AsProvided
        pAsRequired = AsRequired
    End If
    
    CalculateAreaSteelFactor

End Sub

'------------------------------------------------------------------------------
' DESCRIPTION:  Calculates modification factor that increases development
'               length and stores it in pTotalIncreaseFactor
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/16/2016   Created
'
'------------------------------------------------------------------------------
Private Sub CalculateIncreaseModificationFactor()

    Const EPOXYANDTOPLIMIT As Double = 1.7
    
    pTotalIncreaseFactor = 1#   ' base factor with no modification
    
    ' check for each flag and multiply to factor variable
    
    ' check for top steel
    If p12_InchesBelowBar = True Then
        pTotalIncreaseFactor = pTotalIncreaseFactor * BAR12INCHESFRESHBELOW
    End If
    
    ' check epoxy next and impose limit of 1.7 to combination of top
    ' reinforcement factor with epoxy coating
    Dim epoxyFactor As Double
    epoxyFactor = 1#
    
    If pEpoxyCoatedCoverAndSpacing = True Then
        epoxyFactor = EPOXYCOATEDCOVERORSPACING
    ElseIf pEpoxyCoatedOther = True Then
        epoxyFactor = EPOXYCOATEDCATCHALL
    End If
    
    If (pTotalIncreaseFactor * epoxyFactor) > EPOXYANDTOPLIMIT Then
        pTotalIncreaseFactor = EPOXYANDTOPLIMIT
    Else
        pTotalIncreaseFactor = pTotalIncreaseFactor * epoxyFactor
    End If
    
    ' lightweight concrete with Fct given factor
    If pFctLWConcrete = True Then
        pTotalIncreaseFactor = pTotalIncreaseFactor * pFctFactor
    End If
    
    ' lightweight concrete factor
    If pLWConcrete = True Then
        pTotalIncreaseFactor = pTotalIncreaseFactor _
            * LIGHTWEIGHTCONCRETENOFCT
    End If
    
    'sand lightweight concrete factor
    If pSandLWConcrete = True Then
        pTotalIncreaseFactor = pTotalIncreaseFactor _
            * SANDLIGHTWEIGHTCONCRETENOFCT
    End If
    
End Sub

'------------------------------------------------------------------------------
' DESCRIPTION:  Calclate Factor for lightweight aggregate concrete where fct
'               is specified and store in pFctFactor variable
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/16/2016   Created
'
'------------------------------------------------------------------------------
Private Sub CalculateL_W_GivenFct()

    If pLightWeightFct <= 0 Then
        pFctFactor = 1#
        Exit Sub
    End If
    
    Dim Calc As Double
    Calc = (0.22 * Math.Sqr(pConcreteCompStrength)) / pLightWeightFct
    
    If Calc <= 1# Then
        pFctFactor = 1#
    Else
        pFctFactor = Calc
    End If
    
End Sub

'------------------------------------------------------------------------------
' DESCRIPTION:  Calculates modification factor that decreases development
'               length and stores it in pTotalDecreaseFactor
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/16/2016   Created
'
'------------------------------------------------------------------------------
Private Sub CalculateDecreaseModificationFactor()

    pTotalDecreaseFactor = 1# ' base factor with no modification
    
    ' lateral spacing and cover
    If pCoverAndSpacing = True Then
        pTotalDecreaseFactor = pTotalDecreaseFactor * COVERANDLATERALSPACING
    End If
    
    ' AsReq / AsProvided
    If pAreaSteelUsed = True Then
        If pAsRequired > 0 And pAsProvided > 0 Then
            pTotalDecreaseFactor = pTotalDecreaseFactor * pAreaSteelFactor
        End If
    End If
    
    ' enclosed in spiral
    If pEnclosedInSpiral = True Then
        pTotalDecreaseFactor = pTotalDecreaseFactor * ENCLOSEDINSPIRALFACTOR
    End If
    
End Sub

'------------------------------------------------------------------------------
' DESCRIPTION:  Calculates area of steel modification factor and stores it
'               in pAreaSteelFactor
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/20/2016   Created
'
'------------------------------------------------------------------------------
Private Sub CalculateAreaSteelFactor()

    If pAsRequired > 0 And pAsProvided > 0 Then
        pAreaSteelFactor = (pAsRequired / pAsProvided)
    Else
        pAreaSteelFactor = 1#
    End If

End Sub

'------------------------------------------------------------------------------
' DESCRIPTION:  Called in class destructor and error handling routines to
'               delete objects to avoid memory leak
'
' AUTHOR        DATE        NOTES
'------------------------------------------------------------------------------
' Tom Gerard    6/15/2016   Created
'
'------------------------------------------------------------------------------
Private Sub KillObjects()
    Set pBar = Nothing
End Sub
