Attribute VB_Name = "modBezierCurves"
Option Explicit

' This code is based on the paper written by Kenny Hoff - 'Derivation
' of Incremental Forward-Difference Algorithm for Cubic Bezier Curves
' using the Taylor Series Expansion for Polynamial Approximation'
' and a similar derivation done by myself for the quadratic bezier
' curve following the same method.  While I'm sure it's been done
' many times, I did it more as an exercise to further my
' understanding of the methods and algorithms herein.  Additional
' optimizations have been made here to the algorithm presented in
' the paper.
'
' The paper can be found at:
'  http://www.cs.unc.edu/~hoff/projects/comp236/curves/papers/forwardiff.htm
'
' The Forward-Difference Algorithm is much quicker when calculating
' multiple points on the bezier curve because after calculating
' the starting point and the starting differentials only a few
' additions are required to calculate each additional point on
' the curve.
'
' Cubic Bezier:
'    f(t) = (1-t)^3*A + 3*t*(1-t)^2*B + 3*t^3*(1-t) + t^3*D
'
' Quadratic Bezier:
'    f(t) = (1-t)^2*A + 2*t*(1-t)*B + t^2*C
'
'
' Author: Tim Arheit (tarheit@wcoil.com) (8-18-01)
'

' Given the three points defining a quadratic bezier curve
' (X1,Y1), (X2,Y2) and (X3,Y3) this function calculates the
' Number of points along the curve and places them in the
' preallocated arrays x() and y() starting at position count.
' count is incremented to the last element added.
Public Sub QuadraticBezier(x() As Double, y() As Double, _
   count As Long, ByVal Number As Long, _
   ByVal X1 As Double, ByVal Y1 As Double, _
   ByVal x2 As Double, ByVal y2 As Double, _
   ByVal x3 As Double, ByVal y3 As Double)

   Dim f As Double
   Dim df As Double
   Dim ddf As Double
   Dim dt As Double
   
   Dim dTemp1 As Double
   Dim dTemp2 As Double
   Dim dTemp3 As Double
   
   Dim i As Long
   Dim lb As Long
   Dim ub As Long
   
   dt = 1# / CDbl(Number)
   dTemp1 = dt * dt
   lb = count
   ub = count + Number - 1
   
   ' calculate all x coordinates.
   '
   ' f(0) = A
   ' df(0) = (2*B-2*A)*dt + (A-2*B+C)*dt^2
   ' ddf(0) = 2*(A-2*B+C)*dt^2
   '
   dTemp2 = 2 * x2
   dTemp3 = (X1 - dTemp2 + x3) * dTemp1
   
   f = X1
   df = (dTemp2 - 2 * X1) * dt + dTemp3
   ddf = 2 * dTemp3
   
   
   For i = lb To ub
      x(i) = f
      f = f + df
      df = df + ddf
   Next i
   
   
   ' Calculate all y coordinates
   dTemp2 = 2 * y2
   dTemp3 = (Y1 - dTemp2 + y3) * dTemp1
   
   f = Y1
   df = (dTemp2 - 2 * Y1) * dt + dTemp3
   ddf = 2 * dTemp3
   
   For i = lb To ub
      y(i) = f
      f = f + df
      df = df + ddf
   Next i
   
   count = ub + 1
End Sub

Public Sub CubicBezier(x() As Double, y() As Double, _
   count As Long, ByVal Number As Long, _
   ByVal X1 As Double, ByVal Y1 As Double, _
   ByVal x2 As Double, ByVal y2 As Double, _
   ByVal x3 As Double, ByVal y3 As Double, _
   ByVal X4 As Double, ByVal Y4 As Double)
      
   Dim f As Double
   Dim df As Double
   Dim ddf As Double
   Dim dddf As Double
   Dim dt As Double
   
   Dim dTempP1 As Double
   Dim dTempP2 As Double
   Dim dTempP3 As Double
   Dim dTempP4 As Double
   Dim dTempP5 As Double
   
   Dim dTemp1 As Double
   Dim dTemp2 As Double
   
   Dim i As Long
   Dim ub As Long
   Dim lb As Long
   
   
   dt = 1# / CDbl(Number)
   lb = count
   ub = count + Number - 1
   
   dTempP1 = 3 * dt
   dTempP2 = dTempP1 * dt
   dTempP3 = dt ^ 3
   dTempP4 = dTempP2 * 2
   dTempP5 = dTempP3 * 6
   
   ' Calculate x
   '  f(0) = A
   '  df(0) = (B-A)*(3*dt) + (A-2*B+C)*(3*dt^2) + (3*(B-C)-A+D)*(dt^3)
   '  ddf(0) = (A-2*B+C)*(6*dt^2) + (3*(B-C)-A+D)*(6*dt^3)
   '  dddf(0) = (3*(B-C)-A+D)*(6*dt^3)
   
   dTemp1 = (X1 - 2 * x2 + x3)
   dTemp2 = (3 * (x2 - x3) - X1 + X4)
   
   f = X1
   df = (x2 - X1) * dTempP1 + dTemp1 * dTempP2 + dTemp2 * dTempP3
   dddf = dTemp2 * dTempP5
   ddf = dTemp1 * dTempP4 + dddf
   
   For i = lb To ub
      x(i) = f
      f = f + df
      df = df + ddf
      ddf = ddf + dddf
   Next i
   
   
   ' Calculate y
   dTemp1 = (Y1 - 2 * y2 + y3)
   dTemp2 = (3 * (y2 - y3) - Y1 + Y4)
   
   f = Y1
   df = (y2 - Y1) * dTempP1 + dTemp1 * dTempP2 + dTemp2 * dTempP3
   dddf = dTemp2 * dTempP5
   ddf = dTemp1 * dTempP4 + dddf
   
   For i = lb To ub
      y(i) = f
      f = f + df
      df = df + ddf
      ddf = ddf + dddf
   Next i
   
   count = ub + 1
End Sub

