Attribute VB_Name = "modFIXED"
Option Explicit

'
' General support for the FIXED data type used by the API
'
' Send all bugs, improvements, suggestions, etc to the author:
'    Tim Arheit
'    tarheit@wcoil.com
'
' History:
'   08-14-00: Completed.
'

Public Type FIXED
   Fract As Integer
   Value As Integer
End Type

' Converts a FIXED type that is stored in a Long to a Double
' This is likely the fastest of the conversion functions and probably
' should be used if at all possible.
Public Function DoubleFromFixedAsLong(ByVal i As Long) As Double
   DoubleFromFixedAsLong = CDbl(i) / 65536#
End Function

Public Function FixedFromDouble(ByVal d As Double) As FIXED
   Dim f As FIXED
   Dim i As Long
   
   ' Calculate the Value portion
   ' Note: -1.2 must be rounded to -2,  The Value portion can be
   ' positive or negative, but the Fract portion can only be
   ' positive.  Hence -1.2 is stored as -2 + 0.8
   i = Int(d)
   If i < 0 Then
      f.Value = &H8000 Or CInt(i And &H7FFF)
   Else
      f.Value = CInt(i And &H7FFF)
   End If
      
   i = (CLng(d * 65536#) And &HFFFF&)
   If (i And &H8000&) = &H8000& Then
      f.Fract = &H8000 Or CInt(i And &H7FFF&)
   Else
      f.Fract = CInt(i And &H7FFF&)
   End If
   
   FixedFromDouble = f
End Function

Public Function DoubleFromFixed(f As FIXED) As Double
   Dim d As Double
   
   d = CDbl(f.Value)
   
   If f.Fract < 0 Then
      d = d + (32768 + (f.Fract And &H7FFF)) / 65536#
   Else
      d = d + CDbl(f.Fract) / 65536#
   End If
   
   DoubleFromFixed = d
End Function

'
' Returns: (f1+f2)/2
' This is only implemented because it's used a lot in some code and it
' is about twice as fast as the following, plus you don't have the loss
' due to rounding at with the extra unit conversions.
'   FixedFromDouble((DoubleFromfixed(f1) + DoubleFromFixed(f2)) / 2#)
'
Public Function AverageFixed(f1 As FIXED, f2 As FIXED) As FIXED
   Dim f As FIXED
   Dim iFract As Long
   Dim iValue As Long
   
   iValue = CLng(f1.Value) + CLng(f2.Value)
   
   ' Determine if there will be a remainder from
   ' division and add it to iFract
   If (iValue And &H1&) = &H1 Then
      If iValue < 0 Then
         iFract = -&H10000
      Else
         iFract = &H10000
      End If
   Else
      iFract = 0
   End If
   
   ' divide the value
   iValue = iValue \ 2&
      
      
   ' Add the unsigned Fract parts
   If f1.Fract < 0 Then
      iFract = iFract + &H8000& + CLng(f1.Fract And &H7FFF)
   Else
      iFract = iFract + CLng(f1.Fract)
   End If
   
   If f2.Fract < 0 Then
      iFract = iFract + &H8000& + CLng(f2.Fract And &H7FFF)
   Else
      iFract = iFract + CLng(f2.Fract)
   End If
   
   ' Divide the fractional part
   iFract = iFract \ 2&
   
   
   ' Handle overflow and underflow
   If iFract < 0 Then
      iValue = iValue - 1
      iFract = &HFFFF& + iFract
   ElseIf iFract > &HFFFF& Then
      iValue = iValue + 1
      iFract = iFract - &HFFFF&
   End If
   
   
   ' Set the Value
   f.Value = CInt(iValue)
   
   ' Set the fract
   If (iFract And &H8000&) = &H8000& Then
      f.Fract = &H8000 Or CInt(iFract And &H7FFF&)
   Else
      f.Fract = CInt(iFract)
   End If
   
   ' Return the value
   AverageFixed = f
End Function
