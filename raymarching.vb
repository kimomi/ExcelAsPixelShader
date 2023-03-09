Option Explicit

Function GetGrayShade(x As Double, y As Double) As Double
    GetGrayShade = 0
    Dim startPos As Variant
    startPos = Array(0, 0, -5)
    Dim direction As Variant
    direction = Normalize(Array(x - startPos(0), y - startPos(1), 0# - startPos(2)))
    Dim maxDistance As Double
    maxDistance = 20

    Dim distance As Double
    distance = Raymarch(startPos, direction, maxDistance)
    If distance > 0 Then
        GetGrayShade = Render(Array(startPos(0) + direction(0) * distance, startPos(1) + direction(1) * distance, startPos(2) + direction(2) * distance), startPos)
    End If
End Function

Function Raymarch(startPoint As Variant, direction As Variant, maxDistance As Double) As Double
    Const EPSILON As Double = 0.0001
    
    Dim distance As Double
    distance = EPSILON
    
    Dim currentPos As Variant
    currentPos = startPoint
    
    While distance < maxDistance
        Dim distToSphere As Double
        distToSphere = SDFMAP(currentPos)
        
        If distToSphere < EPSILON Then
            Raymarch = distance
            Exit Function
        End If
        
        currentPos = Array(currentPos(0) + direction(0) * distToSphere, currentPos(1) + direction(1) * distToSphere, currentPos(2) + direction(2) * distToSphere)
        distance = distance + distToSphere
    Wend
    
    Raymarch = 0
End Function

Function SDFMAP(pos As Variant) As Double
    SDFMAP = SphereSDF(pos, Array(0, 0, 0), 0.5)

    Dim planeSDFValue As Double
    planeSDFValue = PlaneSDF(pos)
    If SDFMAP > planeSDFValue Then
        SDFMAP = planeSDFValue
    End If
End Function

Function SphereSDF(point As Variant, center As Variant, radius As Double) As Double
    SphereSDF = POINT_DISTANCE(center, point) - radius
End Function

Function POINT_DISTANCE(p1 As Variant, p2 As Variant) As Double
    POINT_DISTANCE = Sqr((p2(0) - p1(0)) ^ 2 + (p2(1) - p1(1)) ^ 2 + (p2(2) - p1(2)) ^ 2)
End Function

Function PlaneSDF(p As Variant) As Double
    PlaneSDF = p(1) + 0.5
End Function

Function NORMALE_SDF(p As Variant) As Variant
    Dim a As Variant
    Dim b As Variant
    Dim c As Variant
    Dim d As Variant
    Dim PERCISION As Double
    PERCISION = 0.01
    a = VEC_MUL_NUM(Array(1, -1, -1), SDFMAP(VEC_ADD(p, Array(PERCISION, -PERCISION, -PERCISION))))
    b = VEC_MUL_NUM(Array(-1, -1, 1), SDFMAP(VEC_ADD(p, Array(-PERCISION, -PERCISION, PERCISION))))
    c = VEC_MUL_NUM(Array(-1, 1, -1), SDFMAP(VEC_ADD(p, Array(-PERCISION, PERCISION, -PERCISION))))
    d = VEC_MUL_NUM(Array(1, 1, 1), SDFMAP(VEC_ADD(p, Array(PERCISION, PERCISION, PERCISION))))
    NORMALE_SDF = Normalize(VEC_ADD(VEC_ADD(a, b), VEC_ADD(c, d)))
End Function

Function VEC_ADD(a As Variant, b As Variant) As Variant
    VEC_ADD = Array(a(0) + b(0), a(1) + b(1), a(2) + b(2))
End Function

Function VEC_SUB(a As Variant, b As Variant) As Variant
    VEC_SUB = Array(a(0) - b(0), a(1) - b(1), a(2) - b(2))
End Function

Function VEC_MUL(a As Variant, b As Variant) As Variant
    VEC_MUL = Array(a(0) * b(0), a(1) * b(1), a(2) * b(2))
End Function

Function VEC_DOT(a As Variant, b As Variant) As Variant
    VEC_DOT = a(0) * b(0) + a(1) * b(1) + a(2) * b(2)
End Function

Function VEC_MUL_NUM(a As Variant, b As Double) As Variant
    VEC_MUL_NUM = Array(a(0) * b, a(1) * b, a(2) * b)
End Function

Function Normalize(vec As Variant) As Variant
    Dim a As Double
    a = Sqr(vec(0) ^ 2 + vec(1) ^ 2 + vec(2) ^ 2)
    
    Normalize = Array(vec(0) / a, vec(1) / a, vec(2) / a)
End Function

Function Render(p As Variant, ro As Variant) As Double
    Render = 0.2

    Dim n As Variant
    n = NORMALE_SDF(p)
    Dim lightPos As Variant
    lightPos = Array(1.5, 2#, 0#)
    
    Dim shadow As Double
    shadow = SoftShadow(p, lightPos)

    Dim diff As Double
    diff = CLAMP(VEC_DOT(Normalize(VEC_SUB(lightPos, p)), n), 0#, 1#)
    Render = Render + 0.3 * diff * shadow

    Dim specular As Double
    Dim h As Variant
    h = Normalize(VEC_ADD(Normalize(VEC_SUB(ro, p)), Normalize(VEC_SUB(lightPos, p))))
    Render = Render + 0.5 * (CLAMP(VEC_DOT(h, n), 0, 1) ^ 50) * shadow
End Function

Function CLAMP(x As Double, lower As Double, upper As Double) As Double
    If x < lower Then
        CLAMP = lower
    ElseIf x > upper Then
        CLAMP = upper
    Else
        CLAMP = x
    End If
End Function


Function SoftShadow(ro As Variant, rd As Variant) As Double
    SoftShadow = 1#
    Dim t As Double
    t = 0.1
    While t < 8#
        Dim h As Double
        h = SDFMAP(VEC_ADD(ro, VEC_MUL_NUM(rd, t)))
        If (h < 0.001) Then
            SoftShadow = 0#
            Exit Function
        End If
        SoftShadow = WorksheetFunction.Min(SoftShadow, 8# * h / t)
        t = t + h
    Wend
End Function

